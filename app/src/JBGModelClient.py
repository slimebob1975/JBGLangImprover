"""
Gemensam skapelse av OpenAI-klienten.

OpenAI-biblioteket försöker själv igen vid hastighetsbegränsning (429),
serverfel, tidsgränser och avbrutna anslutningar, med växande väntetid och
med hänsyn till serverns Retry-After. Därför behövs ingen fast paus mellan
anropen; här anges bara hur många nya försök som görs innan ett anrop räknas
som misslyckat.
"""

import openai

MODEL_CALL_MAX_RETRIES = 4


def create_openai_client(api_key: str):
    return openai.OpenAI(api_key=api_key, max_retries=MODEL_CALL_MAX_RETRIES)
