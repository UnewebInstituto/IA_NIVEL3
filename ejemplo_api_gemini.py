# AIzaSyDcE_FCAaW71vBbZ4QmhwNphV0WMgGd2dU

from google import genai

client = genai.Client(api_key="AIzaSyDcE_FCAaW71vBbZ4QmhwNphV0WMgGd2dU")

response = client.models.generate_content(
    model="gemini-3-flash-preview",
    contents="¿Las api key generadas por google que son gratuitas tienen tiempo de expiración y las que son pagas, por cuánto tiempo puede usarse?",
)

print(response.text)

