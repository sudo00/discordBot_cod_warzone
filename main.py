import discord
import xlsxwriter
import random
import sqlite3
import pandas as pd
from discord import File
from config import *
import os
import time
import requests


def isAdmin(user_id):
    return (
        (user_id == ADMIN1)
        |
        (user_id == ADMIN2)
        |
        (user_id == ADMIN3)
    )


def isRegStatusOn(status):
    sql = "SELECT status FROM regstat where type = " + str(status)
    cursor.execute(sql)
    status = cursor.fetchone()[0]
    return status


class MyClient(discord.Client):
    async def on_ready(self):
        await self.change_presence(status=discord.Status.online, activity=discord.Activity(type=discord.ActivityType.playing, name="!start - ЛС"))
        return

    async def clearChat(channel):
        async for message in channel.history(limit=5):
            if message.author.id == BOTID:
                await message.delete()

    async def on_raw_reaction_add(self, payload):
        # Выбор режима
        if (payload.user_id != BOTID):
            channelid = payload.channel_id
            userid = payload.user_id
            channel = await self.fetch_channel(channelid)
            async for message in channel.history(limit=5):
                if message.author.id == BOTID:
                    await message.delete()

            if str(payload.emoji) == "1️⃣":
                gameType = 'solo'
                if (isRegStatusOn(1)):
                    msg = await message.channel.send('Введите никнейм игрока.\n\nНапример: "Игрок1".')
                else:
                    msg = await message.channel.send('Регистрация на турнир SOLO закрыт! Чтобы попробовать еще раз или зарегистрироваться на другой турнир напишите "!start"')
                    return False
            if str(payload.emoji) == "2️⃣":
                gameType = 'duo'
                if (isRegStatusOn(2)):
                    msg = await message.channel.send('Введите никнеймы двух игроков через запятую.\n\nНапример: "Игрок1, Игрок2".')
                else:
                    msg = await message.channel.send('Регистрация на турнир DUO закрыт! Чтобы попробовать еще раз или зарегистрироваться на другой турнир напишите "!start"')
                    return False
            if str(payload.emoji) == "3️⃣":
                gameType = 'trio'
                if (isRegStatusOn(3)):
                    msg = await message.channel.send('Введите никнеймы трех игроков через запятую.\n\nНапример: "Игрок1, Игрок2, Игрок3".')
                else:
                    msg = await message.channel.send('Регистрация на турнир TRIO закрыт! Чтобы попробовать еще раз или зарегистрироваться на другой турнир напишите "!start"')
                    return False
            if str(payload.emoji) == "4️⃣":
                gameType = 'squad'
                if (isRegStatusOn(4)):
                    msg = await message.channel.send('Введите никнеймы четырех игроков через запятую.\n\nНапример: "Игрок1, Игрок2, Игрок3, Игрок4".')
                else:
                    msg = await message.channel.send('Регистрация на турнир SQUAD закрыт! Чтобы попробовать еще раз или зарегистрироваться на другой турнир напишите "!start"')
                    return False
            if str(payload.emoji) == "☑️":
                sql = "SELECT max(updated_at) FROM team WHERE user_id='" + \
                    str(userid) + "'"
                cursor.execute(sql)
                updated_at = cursor.fetchone()[0]
                sql = "SELECT type FROM team WHERE user_id='" + \
                    str(userid) + \
                    "' and updated_at = " + str(updated_at)
                cursor.execute(sql)
                result = cursor.fetchone()
                gameType = result[0]               
                if gameType == 'solo':
                    payType = SOLO_PAY
                    payUrl = SOLO_PAY_URL
                elif gameType == 'duo':
                    payType = DUO_PAY
                    payUrl = DUO_PAY_URL
                elif gameType == 'trio':
                    payType = TRIO_PAY
                    payUrl = TRIO_PAY_URL
                elif gameType == 'squad':
                    payType = SQUAD_PAY
                    payUrl = SQUAD_PAY_URL

                if payType:
                    sql = "SELECT `check` FROM team WHERE user_id='" + \
                        str(userid) + \
                        "' and updated_at = " + str(updated_at)
                    cursor.execute(sql)
                    result = cursor.fetchone()
                    checkVal = result[0]        
                    response = requests.post('https://yoomoney.ru/api/operation-history', data={'details':'true', 'type':'deposition'}, headers={'Authorization': 'Bearer ' + str(YOOMONEY_TOKEN)})
                    json_response = response.json()['operations']
                    payExist = False
                    for key in json_response:
                        if 'message' in key.keys():
                            if str(checkVal) in key['message']:
                                payExist = True
                                break
                    if payExist == False:                        
                        reactions = ["☑️"]
                        msg = await message.channel.send('Платеж не найден! Для внесения взноса перейдите по ссылке - ' + str(payUrl) + '. ВАЖНО! В комментариях к оплате укажите код подтверждения вашей команды - ' + str(checkVal) + '.\n'
                            'После проведения оплаты нажмите галку для подтверждения регистрации. В случае, если вы провели оплату, но вы видите эту ошибку - свяжитесь с администрацией.')
                        for name in reactions:
                            await msg.add_reaction(name)
                if payType == False or (payType == True and payExist == True):
                    sql = "SELECT number FROM team WHERE type='" + \
                        gameType + "' ORDER BY number"
                    cursor.execute(sql)
                    result = cursor.fetchall()

                    if len(result) != 0:
                        maxTeam = result[len(result) - 1][0]
                        inputInPass = False
                        for currnumber in range(int(maxTeam)):
                            if currnumber == 0:
                                continue
                            exist = False
                            for teamNumber in result:
                                if int(teamNumber[0]) == currnumber:
                                    exist = True
                            if exist:
                                continue
                            else:
                                inputInPass = True
                                number = currnumber
                                break

                        if inputInPass != True:
                            number = int(
                                result[len(result) - 1][0]) + 1
                    else:
                        number = 1

                    userData = [
                        (str(number), str(userid), str(gameType))]
                    cursor.executemany(
                        "UPDATE team SET number = ? WHERE user_id = ? and type = ?", userData)
                    userData = [
                        ('0;0;0;0;0', str(userid), str(gameType))]
                    cursor.executemany(
                        "UPDATE team SET points = ? WHERE user_id = ? and type = ?", userData)
                    conn.commit()
                    msg = await message.channel.send('Готово! Ваша команда зарегистрированна на турнир ' + str(gameType).upper() + ' под №' + str(number) + '.\n\nЕсли хотите пересоздать команду, напишите "!start".')
                    if str(gameType).upper() == 'SOLO':
                        giveRole = SOLOROLE
                    elif str(gameType).upper() == 'DUO':
                        giveRole = DUOROLE
                    elif str(gameType).upper() == 'TRIO':
                        giveRole = TRIOROLE
                    elif str(gameType).upper() == 'SQUAD':
                        giveRole = SQUADROLE
                    server = self.get_guild(id=SERVER)
                    role = discord.utils.get(server.roles, id=giveRole)
                    member = await server.fetch_member(userid)
                    await member.add_roles(role)
                    return True
            
            sql = "SELECT * FROM team WHERE user_id=? and type ='" + \
                str(gameType) + "'"
            cursor.execute(sql, [(str(userid))])
            if (cursor.fetchone() is None):
                cursor.execute("INSERT INTO team VALUES (" + str(userid) +
                               ",0,'','','','" + str(gameType) + "','0;0;0;0;0', " + str(time.time()) + ", 0)")
            else:
                cursor.execute("UPDATE team SET updated_at = " + str(time.time()) +
                               " WHERE user_id = '" + str(userid) + "' and type = '" + str(gameType) + "'")
            conn.commit()

    async def on_message(self, message):
        if (message.author.id != BOTID):
            # Сохранение результатов SOLO
            if (message.channel.id == SOLOCHANNEL):
                result = await savePoints(message, cursor, 'solo')
            # Сохранение результатов DUO
            if (message.channel.id == DUOCHANNEL):
                result = await savePoints(message, cursor, 'duo')
            # Сохранение результатов TRIO
            if (message.channel.id == TRIOCHANNEL):
                result = await savePoints(message, cursor, 'trio')
            # Сохранение результатов SQUAD
            if (message.channel.id == SQUADCHANNEL):
                result = await savePoints(message, cursor, 'squad')

            # Предложение регистрации
            if ((str(message.channel.type) == 'private') & (str(message.content) != '!start')):
                sql = "SELECT * FROM team"
                cursor.execute(sql)
                async for msg in message.channel.history(limit=2):
                    if msg.author.id == BOTID:
                        if (msg.content.find('Введите никнейм') == 0):
                            async for msgDel in message.channel.history(limit=5):
                                if msgDel.author.id == BOTID:
                                    await msgDel.delete()
                            gamers = str(message.content)
                            sql = "SELECT max(updated_at) FROM team WHERE user_id='" + \
                                str(message.author.id) + "'"
                            cursor.execute(sql)
                            updated_at = cursor.fetchone()[0]
                            userData = [(0, str(message.author.id))]
                            cursor.executemany(
                                "UPDATE team SET number = ? WHERE user_id = ? and updated_at = " + str(updated_at), userData)
                            cursor.execute("UPDATE team SET gamers = '" + gamers +
                                           "' WHERE user_id = '" + str(message.author.id) + "' and updated_at = " + str(updated_at))
                            conn.commit()
                            msg = await message.channel.send('Введите платформы игроков.\n\nНапример: "PC, PS4, XBOX".')
                            break
                        if (msg.content.find('Введите платформы') == 0):
                            async for msgDel in message.channel.history(limit=5):
                                if msgDel.author.id == BOTID:
                                    await msgDel.delete()
                            platforms = str(message.content)
                            sql = "SELECT max(updated_at) FROM team WHERE user_id='" + \
                                str(message.author.id) + "'"
                            cursor.execute(sql)
                            updated_at = cursor.fetchone()[0]
                            cursor.execute("UPDATE team SET platforms = '" + platforms +
                                           "' WHERE user_id = '" + str(message.author.id) + "' and updated_at = " + str(updated_at))
                            contact = str(message.author)
                            cursor.execute("UPDATE team SET contact = '" + contact +
                                           "' WHERE user_id = '" + str(message.author.id) + "' and updated_at = " + str(updated_at))
                            conn.commit()
                            sql = "SELECT type FROM team WHERE user_id='" + \
                                str(message.author.id) + \
                                "' and updated_at = " + str(updated_at)
                            cursor.execute(sql)
                            result = cursor.fetchone()
                            gameType = result[0]

                            if gameType == 'solo':
                                payType = SOLO_PAY
                                payUrl = SOLO_PAY_URL
                            elif gameType == 'duo':
                                payType = DUO_PAY
                                payUrl = DUO_PAY_URL
                            elif gameType == 'trio':
                                payType = TRIO_PAY
                                payUrl = TRIO_PAY_URL
                            elif gameType == 'squad':
                                payType = SQUAD_PAY
                                payUrl = SQUAD_PAY_URL

                            reactions = ["☑️"]
                            if (payType == False):
                                msg = await message.channel.send('Нажмите на галку для подтверждения регистрации на турнир!')
                            else:
                                checkVal = random.randint(100000,999999)
                                cursor.execute("UPDATE team SET `check` = " + str(checkVal) + " WHERE user_id='" + \
                                str(message.author.id) + \
                                "' and updated_at = " + str(updated_at))
                                conn.commit()
                                msg = await message.channel.send('Внимание! Вы регистрируетесь на турнир c денежным взносом.\n' +
                                'Для внесения взноса перейдите по ссылке - ' + str(payUrl) + '. ВАЖНО! В комментариях к оплате укажите код подтверждения вашей команды - ' + str(checkVal) + '.\n'
                            'После проведения оплаты нажмите галку для подтверждения регистрации.')
                            for name in reactions:
                                await msg.add_reaction(name)
                                


            # !start - Начало регистрации
            if ((str(message.channel.type) == 'private') & (str(message.content) == '!start')):
                reactions = ["1️⃣", "2️⃣", "3️⃣", "4️⃣"]
                msg = await message.channel.send('Выберите турнир: SOLO, DUO, TRIO, SQUAD!')
                for name in reactions:
                    await msg.add_reaction(name)
                return True
            # !save - Загрузка файлов
            if (
                (str(message.channel.type) == 'private')
                & (str(message.content) == '!save')
                & isAdmin(message.author.id)
            ):

                saveExcel('solo')
                saveExcel('duo')
                saveExcel('trio')
                saveExcel('squad')
                await savePointsInExcel(cursor, 'solo')
                await savePointsInExcel(cursor, 'duo')
                await savePointsInExcel(cursor, 'trio')
                await savePointsInExcel(cursor, 'squad')
                msg = await message.channel.send(
                    content="Файл регистрации: ",
                    files=[
                        File(str(os.getcwd()) + '/register/solo.xlsx'),
                        File(str(os.getcwd()) + '/register/duo.xlsx'),
                        File(str(os.getcwd()) + '/register/trio.xlsx'),
                        File(str(os.getcwd()) + '/register/squad.xlsx')
                    ]
                )
                msg = await message.channel.send(
                    content="Файлы результатов матчей: ",
                    files=[
                        File(str(os.getcwd()) + '/score/solo.xlsx'),
                        File(str(os.getcwd()) + '/score/duo.xlsx'),
                        File(str(os.getcwd()) + '/score/trio.xlsx'),
                        File(str(os.getcwd()) + '/score/squad.xlsx')
                    ]
                )
                return True
            # !clear - Чистка всех турниров
            if ((str(message.channel.type) == 'private') & (message.content.find('!clear') == 0)
                    & isAdmin(message.author.id)
                    ):
                gameType = message.content[len('!clear '):]
                if not(gameType == 'solo' or gameType == 'duo' or gameType == 'trio' or gameType == 'squad'):
                    await message.channel.send(':warning: Неправильно введены данные пример: "!clear duo" :warning:')
                    return True
                sql = "DELETE FROM team WHERE type='" + gameType + "'"
                cursor.execute(sql)
                conn.commit()
                server = self.get_guild(id=SERVER)
                if gameType == 'solo':
                    role = discord.utils.get(server.roles, id=SOLOROLE)
                    for user in role.members:
                        await user.remove_roles(role)
                elif gameType == 'duo':
                    role = discord.utils.get(server.roles, id=DUOROLE)
                    for user in role.members:
                        await user.remove_roles(role)
                elif gameType == 'trio':
                    role = discord.utils.get(server.roles, id=TRIOROLE)
                    for user in role.members:
                        await user.remove_roles(role)
                elif gameType == 'squad':
                    role = discord.utils.get(server.roles, id=SQUADROLE)
                    for user in role.members:
                        await user.remove_roles(role)
                await message.channel.send('Все команды и результаты турнира ' + str(gameType) + ' удалены.')
                return True

            # !clear - Чистка в чате
            if ((str(message.channel.type) != 'private') & (str(message.content) == '!clear')
                    & isAdmin(message.author.id)
                    ):

                await discord.TextChannel.purge(message.channel)
                return True

            # !add - добавление команды в турнир
            if ((str(message.channel.type) == 'private') & (message.content.find('!add') == 0)
                    & isAdmin(message.author.id)
                    ):
                result = message.content[len('!add '):]
                result = result.split(';')
                if (len(result)) != 4:
                    await message.channel.send(':warning: Неправильно введены данные пример: "!add solo;Игрок, Игрок2, Игрок3;PC/xbox;89599999" :warning:')
                    return True
                gameType = result[0]
                gamers = result[1]
                platfroms = result[2]
                contact = result[3]
                sql = "SELECT number FROM team WHERE type='" + \
                    gameType + "' ORDER BY number"
                cursor.execute(sql)
                result = cursor.fetchall()

                if len(result) != 0:
                    maxTeam = result[len(result) - 1][0]
                    inputInPass = False
                    for currnumber in range(int(maxTeam)):
                        if currnumber == 0:
                            continue
                        exist = False
                        for teamNumber in result:
                            if int(teamNumber[0]) == currnumber:
                                exist = True
                        if exist:
                            continue
                        else:
                            inputInPass = True
                            number = currnumber
                            break

                    if inputInPass != True:
                        number = int(result[len(result) - 1][0]) + 1
                else:
                    number = 1

                sql = "INSERT INTO team VALUES (?,?,?,?,?,?,'0;0;0;0;0', " + \
                    str(time.time()) + ", 0)"
                userData = [(str(time.time()), str(number), str(
                    gamers), str(platfroms), str(contact), str(gameType))]
                cursor.executemany(sql, userData)
                conn.commit()
                await message.channel.send('Создана команда №' + str(number) + ' в турнир ' + gameType)
                return True
            # !del - удаление команды из турнира
            if ((str(message.channel.type) == 'private') & (message.content.find('!del') == 0)
                    & isAdmin(message.author.id)
                    ):
                result = message.content[len('!del '):]
                result = result.split('/')
                if (len(result)) != 2:
                    await message.channel.send(':warning: Неправильно введены данные пример: "!del solo/1" :warning:')
                    return True
                gameType = result[0]
                number = result[1]
                sql = "DELETE FROM team WHERE type='" + gameType + "' and number = " + number
                cursor.execute(sql)
                conn.commit()
                await message.channel.send('Команда №' + str(number) + ' удалена из турнира ' + gameType)
                return True
            # !reg - управление регистрацией
            if ((str(message.channel.type) == 'private') & (message.content.find('!reg') == 0)
                    & isAdmin(message.author.id)
                    ):
                result = message.content[len('!reg '):]
                result = result.split('/')
                if (len(result)) != 2:
                    await message.channel.send(':warning: Неправильно введены данные пример: "!reg solo/on" - открыть или "!reg solo/off" - закрыть :warning:')
                    return True

                if result[0] == 'solo':
                    gameType = 1
                elif result[0] == 'duo':
                    gameType = 2
                elif result[0] == 'trio':
                    gameType = 3
                elif result[0] == 'squad':
                    gameType = 4
                else:
                    await message.channel.send(':warning: Такого турнира нет. Есть: solo, duo, trio, squad. :warning:')
                    return False

                if result[1] == 'on':
                    status = 1
                    statusText = ' открыта.'
                elif result[1] == 'off':
                    status = 0
                    statusText = ' закрыта.'
                else:
                    await message.channel.send(':warning: Такого режима нет. Есть: on, off. :warning:')
                    return False

                cursor.execute("UPDATE regstat SET status = " +
                               str(status) + " WHERE type = " + str(gameType))
                conn.commit()
                await message.channel.send('Регистрация на турнир ' + result[0] + statusText)
                return True


async def savePointsInExcel(cursor, gameType):
    sql = "SELECT number, gamers, platforms, points FROM team WHERE type=? and number != 0  ORDER BY number"
    cursor.execute(sql, [(gameType)])
    result = cursor.fetchall()
    number = []
    gamers = []
    platforms = []
    game1 = []
    game2 = []
    game3 = []
    game4 = []
    game5 = []
    score = []
    for team in result:
        number.append(team[0])
        gamers.append(team[1])
        platforms.append(team[2])
        games = team[3].split(';')
        scorePoint = 0
        if (games[0] == ''):
            games[0] = 0
        if 0 < len(games):
            game1.append(games[0])
            scorePoint += int(games[0])
        else:
            game1.append(0)
        if 1 < len(games):
            game2.append(games[1])
            scorePoint += int(games[1])
        else:
            game2.append(0)
        if 2 < len(games):
            game3.append(games[2])
            scorePoint += int(games[2])
        if 3 < len(games):
            game4.append(games[3])
            scorePoint += int(games[3])
        if 4 < len(games):
            game5.append(games[4])
            scorePoint += int(games[4])
        score.append(scorePoint)
    df = pd.DataFrame({
        'Номер команды ': number,
        'Игрок(и)': gamers,
        'Платформа': platforms,
        'Игра 1': game1,
        'Игра 2': game2,
        'Игра 3': game3,
        'Игра 4': game4,
        'Игра 5': game5,
        'Итог': score
    })
    writer = pd.ExcelWriter('./score/' + gameType +
                            '.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name=gameType, index=False)  # send df to writer
    worksheet = writer.sheets[gameType]  # pull worksheet object
    worksheet.set_column(0, 0, 30)  # set column width
    worksheet.set_column(0, 1, 30)  # set column width
    worksheet.set_column(0, 2, 30)  # set column width
    worksheet.set_column(0, 3, 30)  # set column width
    worksheet.set_column(0, 4, 30)  # set column width
    writer.save()


async def savePoints(message, cursor, gameType):
    msg = str(message.content).split('/')
    if (len(msg) != 2):
        await message.channel.send(':warning: <@' + str(message.author.id) + '>, неправильно введены данные! Введите №Игры/Очки. :warning:\n\nНапример (без кавычек) - "1/25".')
        return False

    # Проверка на правильно введенный номер игры
    if int(msg[0]) == 1:
        gameNumber = int(msg[0]) - 1
    elif int(msg[0]) == 2:
        gameNumber = int(msg[0]) - 1
    elif int(msg[0]) == 3:
        gameNumber = int(msg[0]) - 1
    elif int(msg[0]) == 4:
        gameNumber = int(msg[0]) - 1
    elif int(msg[0]) == 5:
        gameNumber = int(msg[0]) - 1
    else:
        await message.channel.send(':warning: <@' + str(message.author.id) + '>, такого № игры нет. Есть: 1, 2, 3, 4, 5 игры. :warning:')
        return False

    # Проверка числа очков на целочисленость значения
    if str(msg[1]).isdigit() != True:
        await message.channel.send(':warning: <@' + str(message.author.id) + '>, неправильно введено число очков, значение очков должно состоять только из целых чисел! :warning:')
        return False

    # Проверка команды в турнире на существование
    sql = "SELECT points, number FROM team WHERE type = '" + \
        gameType + "' and user_id = ?"
    cursor.execute(sql, [(str(message.author.id))])
    result = cursor.fetchone()
    if result is None:
        await message.channel.send(':warning: <@' + str(message.author.id) + '>, на ваш профиль нет зарегистрированных команд. :warning:')
        return False

    points = str(result[0]).split(';')
    points[gameNumber] = str(msg[1])
    pointsText = ';'.join(points)
    cursor.execute("UPDATE team SET points = '" + str(pointsText) +
                   "' WHERE type = '" + gameType + "' and user_id ='" + str(message.author.id) + "'")
    conn.commit()
    await message.channel.send('<@' + str(message.author.id) + '>' +
                               '\n:video_game: Команда №' + str(result[1]) + ' :video_game:' +
                               '\n:white_check_mark: Присвоен счет ' + msg[1] + ' к игре ' + msg[0] + '! :white_check_mark:' +
                               '\n\n:arrow_down: :100: :arrow_down:\nОбщий счет по играм - ' + str(int(points[0]) + int(points[1]) + int(points[2]) + int(points[3]) + int(points[4])) +
                               '\n 1) ' + str(points[0]) +
                               '\n 2) ' + str(points[1]) +
                               '\n 3) ' + str(points[2]) +
                               '\n 4) ' + str(points[3]) +
                               '\n 5) ' + str(points[4]))
    return True


def saveExcel(typeGame):
    sql = "SELECT number, gamers, platforms, contact FROM team WHERE type=? and number != 0 ORDER BY number"
    cursor.execute(sql, [(typeGame)])
    result = cursor.fetchall()
    number = []
    gamers = []
    platforms = []
    contact = []
    for team in result:
        number.append(team[0])
        gamers.append(team[1])
        platforms.append(team[2])
        contact.append(team[3])
    df = pd.DataFrame({
        'Номер команды ': number,
        'Игрок(и)': gamers,
        'Платформа': platforms,
        'Контакты': contact})
    writer = pd.ExcelWriter('./register/' + typeGame +
                            '.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name=typeGame, index=False)  # send df to writer
    worksheet = writer.sheets[typeGame]  # pull worksheet object
    worksheet.set_column(0, 0, 30)  # set column width
    worksheet.set_column(0, 1, 30)  # set column width
    worksheet.set_column(0, 2, 30)  # set column width
    worksheet.set_column(0, 3, 30)  # set column width
    worksheet.set_column(0, 4, 30)  # set column width
    writer.save()


conn = sqlite3.connect('./sqllite/Chinook_Sqlite.sqlite')
cursor = conn.cursor()
client = MyClient()
client.run(TOKEN)
