"""
Gemensam skapelse av OpenAI-klienten.

OpenAI-biblioteket försöker själv igen vid hastighetsbegränsning (429),
serverfel, tidsgränser och avbrutna anslutningar, med växande väntetid och
med hänsyn till serverns Retry-After. Därför behövs ingen fast paus mellan
anropen; här anges bara hur många nya försök som görs innan ett anrop räknas
som misslyckat.
"""

import logging
import os

import openai

MODEL_CALL_MAX_RETRIES = 4

# Högsta antal samtidiga anrop i den lokala granskningen. Den globala
# granskningen körs dessutom vid sidan av dessa. Kan ändras med
# miljövariabeln JBG_MAX_PARALLEL_MODEL_CALLS; 1 ger anrop i följd.
DEFAULT_MAX_PARALLEL_MODEL_CALLS = 4


def max_parallel_model_calls() -> int:
    raw = os.getenv("JBG_MAX_PARALLEL_MODEL_CALLS", "")
    try:
        value = int(raw) if raw.strip() else DEFAULT_MAX_PARALLEL_MODEL_CALLS
    except ValueError:
        logging.getLogger(__name__).warning(
            f"Invalid JBG_MAX_PARALLEL_MODEL_CALLS={raw!r}; using {DEFAULT_MAX_PARALLEL_MODEL_CALLS}"
        )
        value = DEFAULT_MAX_PARALLEL_MODEL_CALLS
    return max(1, value)


def create_openai_client(api_key: str):
    return openai.OpenAI(api_key=api_key, max_retries=MODEL_CALL_MAX_RETRIES)
