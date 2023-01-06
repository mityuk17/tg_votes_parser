import datetime
import pandas as pd
import os

import pyrogram.errors.exceptions.all
from pyrogram import Client
from pyrogram.raw import functions


def main():

    app = Client("app", "api-id", "api-hash")
    app.start()
    if not os.path.isdir('Отчёты'):
        os.mkdir('Отчёты')
    while input('Нажмите Enter для работы, напишите exit для выхода:') != 'exit':
        users = list()
        chat_id = input("Введите username чата: ")

        time = input("Время начала(в формате dd.mm.yy HH:MM:SS): ")
        time_start = datetime.datetime.strptime(time, '%d.%m.%y %H:%M:%S')

        time = input("Время конца(в формате dd.mm.yy HH:MM:SS): ")
        time_end = datetime.datetime.strptime(time, '%d.%m.%y %H:%M:%S')
        try:



            chat = app.resolve_peer(chat_id)
            messages = app.get_chat_history(chat_id)
            flag = False
            for message in messages:
                date = message.date
                if str(message.media) == 'MessageMediaType.POLL' and time_start <= date <= time_end:
                    vote_flag = False
                    flag = True
                    message_id = message.id
                    try:
                        g = app.invoke(functions.messages.GetPollVotes(
                            peer=chat,
                            id=message_id,
                            limit=99999
                        ))
                    except pyrogram.errors.exceptions.bad_request_400.PollVoteRequired:
                        app.vote_poll(chat_id, message_id, 0)
                        vote_flag = True
                        g = app.invoke(functions.messages.GetPollVotes(
                            peer=chat,
                            id=message_id,
                            limit=99999
                        ))
                    pool = app.get_messages(chat_id, message_id)
                    text_poll = pool.poll.question
                    for vote in g.votes:
                        if not(vote_flag) or vote.user_id != app.get_users('me').id:
                            user = app.get_users(vote.user_id)
                            for option in pool.poll.options:
                                try:
                                    if int(option.data) + 1 == int(vote.option,) + 1:
                                        text_answer = option.text
                                        vote_option = int(option.data) + 1
                                except:
                                    if int.from_bytes(option.data, 'little') + 1 == int.from_bytes(vote.option,'little') + 1:
                                        text_answer = option.text
                                        vote_option = int.from_bytes(option.data, 'little') + 1
                            users.append({"data_poll": date.strftime("%d.%m.%Y"), "time_poll": date.strftime("%H:%M:%S"),"id": vote.user_id, "username": user.username,
                                          "first_name": user.first_name, "last_name": user.last_name,"text_poll": text_poll,
                                          "text_answer": text_answer})
                            filename = f'poll_{chat_id}_{datetime.datetime.now().strftime("%d-%m-%y_%H-%M-%S")}.xlsx'
        except pyrogram.errors.exceptions as error:
            print('Произошла ошибка при парсинге опросов', error)
        try:
            if users:
                df = pd.DataFrame.from_dict(users)
                df.to_excel(f'Отчёты/{filename}',index=False)
                print(f'Обработка опросов в группе {chat_id} завершена успешно. Данные сохранены в файле {filename} в папке "Отчёты"')
            if not(flag):
                print('В чате не были найдены опросы, удовлетворяющие условиям.')
        except pd.errors as error:
            print('Произошла ошибка при сохрании отчёта в файл', error)



if __name__ == "__main__":
    main()



