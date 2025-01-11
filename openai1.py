import openai
from config import apikey


def generate_response(prompt):
    openai.api_key = apikey
    response = openai.Completion.create(
        model="gpt-3.5-turbo-instruct",
        prompt=prompt,
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    return response["choices"][0]["text"]


user_input = "Write an email to my boss for resignation."
response_text = generate_response(user_input)
print(response_text)
