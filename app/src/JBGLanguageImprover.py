import os
import sys
import json
import logging

from app.src.JBGCommentsRenderer import CommentsRenderer
from app.src.JBGTokenDiffEngine import TokenDiffEngine

try:
    from app.src.JBGDocumentStructureExtractor import DocumentStructureExtractor
    from app.src.JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI
    from app.src.JBGChangePlanner import ChangePlanner
    from app.src.JBGDocxPackage import DocxPackage
    from app.src.JBGSimpleMarkupRenderer import SimpleMarkupRenderer
    from app.src.JBGTrackedChangesRenderer import TrackedChangesRenderer
except ModuleNotFoundError:
    from JBGDocumentStructureExtractor import DocumentStructureExtractor
    from JBGLangImprovSuggestorAI import JBGLangImprovSuggestorAI
    from JBGChangePlanner import ChangePlanner
    from JBGDocxPackage import DocxPackage
    from JBGSimpleMarkupRenderer import SimpleMarkupRenderer
    from JBGTrackedChangesRenderer import TrackedChangesRenderer


class JBGLanguageImprover:
    """
    Ny orchestration-klass för den ombyggda pipelinen.

    Ansvar:
    - extrahera dokumentstruktur
    - köra AI-suggestorn
    - bygga ChangePlan
    - applicera SimpleMarkupRenderer på DOCX
    - spara slutresultat

    Denna version är avsedd för test av den nya DOCX-kedjan.
    """

    def __init__(
        self,
        input_path,
        api_key,
        model,
        prompt_policy,
        temperature,
        include_motivations,
        logger,
        progress_callback=None,
        save_intermediate_json=True,
        docx_mode=None,
        **kwargs,
    ):
        self.input_path = input_path
        self.api_key = api_key
        self.model = model
        self.prompt_policy = prompt_policy
        self.temperature = temperature
        self.include_motivations = include_motivations
        self.logger = logger
        self.progress_callback = progress_callback
        self.save_intermediate_json = save_intermediate_json

        self.docx_mode = docx_mode or "simple"
        self.extra_kwargs = kwargs

        if self.docx_mode not in {"simple", "tracked"}:
            raise ValueError(f"Unsupported docx_mode: {self.docx_mode}")
        
        if kwargs:
            self.logger.warning(f"Unused JBGLanguageImprover kwargs: {list(kwargs.keys())}")

        self.structure_json = input_path.replace(
            os.path.splitext(input_path)[1], "_structure.json"
        )
        self.suggestions_json = input_path.replace(
            os.path.splitext(input_path)[1], "_suggestions.json"
        )
        self.filter_report_json = input_path.replace(
            os.path.splitext(input_path)[1], "_suggestion_filter_report.json"
        )

        self.structure = None
        self.validated_suggestions = []
        self.change_plans = []
        self.render_results = []

    # ------------------------------------------------------------------
    # Publikt API
    # ------------------------------------------------------------------

    def run(self, output_path=None):
        ext = os.path.splitext(self.input_path)[1].lower()

        if ext != ".docx":
            raise ValueError(
                "Den här nya JBGLanguageImprover-versionen stödjer i detta skede bara .docx "
                "för test av den ombyggda Word-pipelinen."
            )

        self._report("Analyserar dokumentets struktur...")
        self.structure = self._extract_structure()

        self._report("Skickar dokumentet till språkmodellen för förslag...")
        self.validated_suggestions = self._generate_suggestions()

        self._report("Bygger ändringsplan...")
        self.change_plans = self._build_change_plans()

        if self.docx_mode == "tracked":
            self._report("Applicerar spåra ändringar i Word-dokumentet...")
        else:
            self._report("Applicerar enkel markup i Word-dokumentet...")

        final_output_path = self._render_docx(output_path=output_path)

        self._report("Klart.")
        return final_output_path

    # ------------------------------------------------------------------
    # Steg 1: Extraktion
    # ------------------------------------------------------------------

    def _extract_structure(self):
        extractor = DocumentStructureExtractor(self.input_path, self.logger)
        structure = extractor.extract()

        if self.save_intermediate_json:
            saved_path = extractor.save_as_json(self.structure_json)
            self.logger.info(f"Saved structure JSON: {saved_path}")

        return structure

    # ------------------------------------------------------------------
    # Steg 2: Suggestor
    # ------------------------------------------------------------------

    def _generate_suggestions(self):
        ai = JBGLangImprovSuggestorAI(
            api_key=self.api_key,
            model=self.model,
            prompt_policy=self.prompt_policy,
            temperature=self.temperature,
            logger=self.logger,
            progress_callback=self.progress_callback,
            strict_validation=True,
            allow_normalized_matches=True,
        )

        # Suggestorn arbetar mot strukturfilen
        if not os.path.exists(self.structure_json):
            with open(self.structure_json, "w", encoding="utf-8") as f:
                json.dump(self.structure, f, indent=2, ensure_ascii=False)

        ai.load_structure(self.structure_json)
        ai.suggest_changes_token_aware_batching()

        if self.save_intermediate_json:
            suggestions_path = ai.save_as_json(self.suggestions_json, use_validated=True)
            filter_report_path = ai.save_filter_report(self.filter_report_json)
            self.logger.info(f"Saved validated suggestions JSON: {suggestions_path}")
            self.logger.info(f"Saved suggestion filter report: {filter_report_path}")

        self.logger.info(f"Validated suggestions count: {len(ai.validated_suggestions)}")
        return ai.validated_suggestions

    # ------------------------------------------------------------------
    # Steg 3: Planning
    # ------------------------------------------------------------------

    def _build_change_plans(self):
        diff_engine = TokenDiffEngine(logger=self.logger)
        planner = ChangePlanner(
            structure=self.structure,
            diff_engine=diff_engine,
            logger=self.logger,
        )

        plans = planner.build_plans(self.validated_suggestions)

        self.logger.info(f"Built change plans: {len(plans)}")

        overlapping = [p for p in plans if "overlapping_change_conflict" in p.notes]
        if overlapping:
            self.logger.warning(
                f"Detected overlapping change plans: {len(overlapping)}"
            )

        return plans

    # ------------------------------------------------------------------
    # Steg 4: Rendering
    # ------------------------------------------------------------------

    def _render_docx(self, output_path=None):
        if output_path is None:
            base, ext = os.path.splitext(self.input_path)
            suffix = "_tracked_changes" if self.docx_mode == "tracked" else "_simple_markup"
            output_path = f"{base}{suffix}{ext}"

        with DocxPackage(self.input_path, self.logger) as pkg:
            if self.docx_mode == "tracked":
                renderer = TrackedChangesRenderer(pkg, self.logger)
            else:
                renderer = SimpleMarkupRenderer(pkg, self.logger)

            self.render_results = renderer.apply_plans(self.change_plans)

            self.comment_results = []
            if self.include_motivations and self.docx_mode == "tracked":
                comments_renderer = CommentsRenderer(pkg, self.logger)
                self.comment_results = comments_renderer.apply_comments_for_results(self.render_results)

                comment_applied = len([r for r in self.comment_results if r.applied])
                comment_failed = len([r for r in self.comment_results if not r.applied])
                self.logger.info(f"Comments applied: {comment_applied}")
                self.logger.info(f"Comments skipped/failed: {comment_failed}")

                for result in self.comment_results:
                    if not result.applied:
                        target = result.plan.target
                        label = f"{target.element_type}:{target.element_id}"
                        self.logger.warning(f"Comment skipped/failed for {label}: {result.message}")

            final_output_path = pkg.save(output_path)

        applied_count = len([r for r in self.render_results if r.applied])
        failed_count = len([r for r in self.render_results if not r.applied])

        self.logger.info(f"Render mode: {self.docx_mode}")
        self.logger.info(f"Render applied: {applied_count}")
        self.logger.info(f"Render skipped/failed: {failed_count}")

        for result in self.render_results:
            if not result.applied:
                target = result.plan.target
                label = f"{target.element_type}:{target.element_id}"
                self.logger.warning(f"Render skipped/failed for {label}: {result.message}")

        return final_output_path

    # ------------------------------------------------------------------
    # Hjälpare
    # ------------------------------------------------------------------

    def _report(self, message: str):
        self.logger.info(message)
        if self.progress_callback is not None:
            try:
                self.progress_callback(message)
            except Exception as ex:
                self.logger.warning(f"Progress callback failed: {ex}")


def main():
    if len(sys.argv) not in (6, 7, 8):
        print(
            f"Usage: python {os.path.basename(__file__)} "
            "<document_path> <api_key> <model> <prompt_policy_file> <custom_addition_file> [docx_mode] [temperature]"
        )
        sys.exit(1)

    input_path = sys.argv[1]
    api_key = sys.argv[2]
    model = sys.argv[3]
    policy_path = sys.argv[4]
    custom_path = sys.argv[5]

    docx_mode = "simple"
    temperature = 0.3

    if len(sys.argv) >= 7:
        arg6 = sys.argv[6]
        if arg6 in {"simple", "tracked"}:
            docx_mode = arg6
        else:
            temperature = float(arg6)

    if len(sys.argv) == 8:
        temperature = float(sys.argv[7])

    with open(policy_path, encoding="utf-8") as f:
        base_prompt = f.read().strip()

    full_prompt = base_prompt
    if custom_path and os.path.exists(custom_path):
        with open(custom_path, encoding="utf-8") as f:
            custom = f.read().strip()
        if custom:
            full_prompt += "\n\n" + custom

    logger = logging.getLogger("jbg-language-improver")
    logger.setLevel(logging.INFO)
    handler = logging.StreamHandler(sys.stdout)
    handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    logger.handlers.clear()
    logger.addHandler(handler)

    improver = JBGLanguageImprover(
        input_path=input_path,
        api_key=api_key,
        model=model,
        prompt_policy=full_prompt,
        temperature=temperature,
        include_motivations=False,
        logger=logger,
        progress_callback=None,
        save_intermediate_json=True,
        docx_mode=docx_mode,
    )

    try:
        final_path = improver.run()
        logger.info(f"Final output saved to: {final_path}")
    except Exception as ex:
        logger.error(f"Pipeline failed: {ex}")
        sys.exit(1)


if __name__ == "__main__":
    main()