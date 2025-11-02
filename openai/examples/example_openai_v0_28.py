# Basic example that demonstrates prompting openai via API in Python.
# API call is openai.ChatCompletion.create(model, messages, temperature)
# Note the openai version below needs openai version <1.0.0
# python -m pip install "openai<1.0.0"  or pip install "openai<1.0.0" or pip install openai==0.28
# this file uses openai version 0.28

import openai
from dotenv import load_dotenv
import os

load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")

def get_completion(prompt, model="gpt-4o"):
    messages = [{"role": "user", "content": prompt}]
    response = openai.ChatCompletion.create(
        model=model,
        messages=messages,
        temperature=0,  # this is the degree of randomness of the model's output
    )
    return response.choices[0].message["content"]


text = input("Enter a text for translation: ")
prompt = (f"Translate the text delimited by triple backticks into Hindi and Russian```{text}```. "
          f"Omit the backticks in my response")
response = get_completion(prompt)
print(response)