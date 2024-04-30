import requests

from config import tg_token, chat_id


def send_message(bot_token, chat_id, message):

    send_message_url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    params = {
        "chat_id": chat_id,
        "text": message
    }
    response = requests.post(send_message_url, data=params, verify=False)
    if response.status_code != 200:
        print(f"Failed to send message: {response.status_code}")
    else:
        print("Message sent successfully!")


if __name__ == "__main__":

    message = "Hello, this is a test message from your bot!"

    send_message(tg_token, chat_id, message)
