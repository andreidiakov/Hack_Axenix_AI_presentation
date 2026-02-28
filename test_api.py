from openai import OpenAI

# ── LLM client ────────────────────────────────────────────────────────────────
client = OpenAI(
    base_url="http://172.28.4.29:8000/v1",
    api_key="dummy",
)

MODEL_NAME = "/model"


def call_llm(system_prompt: str, user_prompt: str, temperature: float = 0.7) -> str:
    """
    Простейший вызов LLM с системным и пользовательским текстом.
    Можно задать температуру для разнообразия ответов.
    """
    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=temperature,
    )
    return response.choices[0].message.content


# ── Тест ───────────────────────────────────────────────────────────────────
system_prompt = "Ты дружелюбный ассистент."
user_prompt = "Привет! Скажи просто 'Привет мир'."

print(call_llm(system_prompt, user_prompt, temperature=0.5))