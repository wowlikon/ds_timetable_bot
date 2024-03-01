from discord import *

from discord.ext import commands
from discord.ext.tasks import loop

from functools import wraps, partial
import asyncio, aiosqlite

from dotenv import load_dotenv
import ujson, datetime, os
import traceback

import lessons, changes


load_dotenv()
temp = os.environ.get('cache')
wBot = commands.Bot(
    command_prefix='!',
    case_insensitive=True,
    intents=Intents.all()
    )

async def init_db():
    async with aiosqlite.connect('bot.db') as db:
        await db.execute('CREATE TABLE IF NOT EXISTS messages (channel_id INTEGER PRIMARY KEY, group_name TEXT, send_time TIMESTAMP NOT NULL)')
        await db.commit()

async def add_channel(channel: int, group: str, h: int, m: int):
    async with aiosqlite.connect('bot.db') as db:
        dt = datetime.datetime.now().replace(hour=h, minute=m)
        await db.execute('INSERT OR REPLACE INTO messages VALUES (?, ?, ?)', (channel, group, dt))
        await db.commit()

async def check_send_time():
    result = []
    current_time = datetime.datetime.now()
    async with aiosqlite.connect('bot.db') as db:
        cursor = await db.execute('SELECT * FROM messages WHERE send_time < ?', (current_time,))
        results = await cursor.fetchall()
        for (ch, group, time) in results:
            result.append((ch, group))
            time = datetime.datetime.strptime(time, '%Y-%m-%d %H:%M:%S.%f')
            new_time = (time + datetime.timedelta(days=1)).strftime('%Y-%m-%d %H:%M:%S.%f')
            await db.execute('UPDATE messages SET send_time = ? WHERE channel_id = ?', (new_time, ch))
            
        await db.commit()
        return result

async def delete_msg_channel(channel: int):
    async with aiosqlite.connect('bot.db') as db:
        await db.execute('DELETE FROM messages WHERE channel_id = ?', (channel,))
        await db.commit()

def async_run(func):
    @wraps(func)
    async def run(*args, loop=None, executor=None, **kwargs):
        if loop is None: loop = asyncio.get_running_loop()
        pfunc = partial(func, *args, **kwargs)
        return await loop.run_in_executor(executor, pfunc)
    return run

@async_run
def parseTTFile(file: str, cache: str) -> str:
    return lessons.parseFile(file, cache)

@async_run
def parseTTCache(cache: str, group: str, day: str) -> lessons.Lessons:
    return lessons.parseCache(cache, group, day)

@async_run
def parseChFile(file: str, cache: str) -> str:
    return changes.parseFile(file, cache)

@async_run
def parseChCache(cache: str, group: str, day: str) -> changes.Changes:
    return changes.parseCache(cache, group, day)

@wBot.event
async def on_ready():
    await init_db()
    print(f"{wBot.user} is ready and online! (ID: {wBot.user.id})") #type: ignore

async def get_days(ctx: AutocompleteContext):
    try:
        with open(temp+"tt_meta_data.json", 'r', encoding='utf-8') as f:
            return ujson.load(f).get('days', {})
        
    except FileNotFoundError as e:
        print(getattr(e, 'message', str(e)))
        return []

async def get_groups(ctx: AutocompleteContext):
    try:
        with open(temp+"tt_meta_data.json", 'r', encoding='utf-8') as f:
            return ujson.load(f).get('groups', {})
        
    except FileNotFoundError as e:
        print(getattr(e, 'message', str(e)))
        return []

@wBot.event
async def on_message(message):
    if message.author == wBot.user: return
    contentType = None
    
    try:
        if message.content == "timetable":
            if len(message.attachments) != 1: raise ValueError("Incorrect file count.")
            
            f = message.attachments[0]
            if f.filename.split(".")[-1] != "xlsx": raise ValueError("Incorrect file type.")
            await f.save(temp+"timetable.xlsx")
            
            embed = Embed(title=f"__**Расписание загружено**__", color=0x4488ee)
            embed.add_field(name=f'**{message.author}**', value=f'> файл успешно сохранён идёт проверка...', inline=False)
            contentType = 'tt'
        
        elif message.content == "changes":
            if len(message.attachments) != 1: raise ValueError("Incorrect file count.")
            
            f = message.attachments[0]
            if f.filename.split(".")[-1] != "xls": raise ValueError("Incorrect file type.")
            await f.save(temp+"changes.xls")
            
            embed = Embed(title=f"__**Изменения загружены**__", color=0x4488ee)
            embed.add_field(name=f'**{message.author}**', value=f'> файл успешно сохранён идёт проверка...', inline=False)
            contentType = 'ch'
        
        else: return
        
        try: await message.delete()
        except: print("Delete failed")
        
        await message.channel.send(message.author.mention, embed=embed)
        
    except Exception:
        embed = Embed(title=f"__**Нужно выбрать 1 файл**__", color=0xee0000)
        embed.add_field(name=f'{message.author}', value=f'> сообщение должно содержать\n> один файл в формате таблицы', inline=False)
        await message.channel.send(message.author.mention, embed=embed)
        
        try: await message.delete()
        except: print("Delete failed")
        
        traceback.print_exc() #INFO
    
    if contentType == 'tt':
        cf = await parseTTFile(temp+f"timetable.xlsx", temp)
        with open(cf+"tt_meta_data.json", 'r', encoding='utf-8') as f:
            metadata = ujson.load(f)
        
        embed = Embed(title=f"__**Расписание проверенно**__", color=0x00ee00)
        
        v = ''
        for i in metadata['groups']: v += f'> {i}\n'
        embed.add_field(name=f'**Группы**',     value=v, inline=False)
        
        v = ''
        for i in metadata['days']: v += f'> {i}\n'
        embed.add_field(name=f'**Дни недели**', value=v,   inline=False)
        await message.channel.send(message.author.mention, embed=embed)
        
    elif contentType == 'ch':
        cf = await parseChFile(temp+"changes.xls", temp)
        with open(cf+"ch_meta_data.json", 'r', encoding='utf-8') as f:
            metadata = ujson.load(f)
        
        embed = Embed(title=f"__**Изменения проверены**__", color=0x00ee00)
        
        v = ''
        for i in metadata['groups']: v += f'> {i}\n'
        embed.add_field(name=f'**Группы**', value=v, inline=False)
        await message.channel.send(message.author.mention, embed=embed)

@wBot.slash_command(name="couples", description="Вывести расписание пар")
async def couples(ctx,
                  group: Option(str, "Введите группу", required=True, autocomplete=utils.basic_autocomplete(get_groups)), # type: ignore
                  day: Option(str, "День недели", required=True, autocomplete=utils.basic_autocomplete(get_days))         # type: ignore
                  ):
    try: await ctx.delete()
    except errors.NotFound: pass
    embed = Embed(title=f"__**Пары на {day} {group}:**__", color=0xffffff)
    coup = await parseTTCache(temp, group, day)
    
    shift = 'Смена неизвестна'
    is_1st = any(map(lambda x: x[1] != "Нету",coup.as_list()[:3]))
    is_2rd = any(map(lambda x: x[1] != "Нету",coup.as_list()[3:]))
    
    if is_1st and not is_2rd: shift = '1 смена'
    if not is_1st and is_2rd: shift = '2 смена'
    
    tmp = coup.as_list()
    for n, les, teach, room in tmp:
        if les != "Нету": info = f'*{teach}* ({room})'
        else: info = '-'
        embed.add_field(name=f'**{n}**', value=f'> **{les}**\n> {info}', inline=False)
    
    try: await message.delete()
    except: print("Delete failed")
    
    if len(tmp):
        await ctx.send(f'**{shift}**', embed=embed)
        await asyncio.sleep(0.05)
        await replaces(ctx, group, day)
    else: await ctx.send('**Расписания нет**')

@wBot.slash_command(name="replaces", description="Вывести замены")
async def replaces(ctx,
                  group: Option(str, "Введите группу", required=True, autocomplete=utils.basic_autocomplete(get_groups)), # type: ignore
                  day: Option(str, "День недели", required=True, autocomplete=utils.basic_autocomplete(get_days))         # type: ignore
                  ):
    try: await ctx.delete()
    except errors.NotFound: pass
    embed = Embed(title=f"__**Замены на {day} {group}:**__", color=0xffffff)
    change = await parseChCache(temp, group, day)
    
    tmp = change.as_list()
    for gr, n, rep, teach, room in tmp:
        embed.add_field(name=f'**{gr}** ({n} пара)', value=f'> **{rep[1]}**\n> ~~{rep[0]}~~\n> *{teach}* ({room})\n\n', inline=False)
    
    try: await message.delete()
    except: print("Delete failed")
    
    if len(tmp): await ctx.send(embed=embed)
    else: await ctx.send('**Замен нет**')

@wBot.slash_command(name="autosend", description="Автоматическая отправка расписания в канал")
async def auto_send(ctx,
                  group: Option(str, "Введите группу", required=True, autocomplete=utils.basic_autocomplete(get_groups)), # type: ignore
                  h: Option(int, "Часов", required=True, ), # type: ignore
                  m: Option(int, "Минут", required=True, ), # type: ignore
                  ):
    try: await ctx.delete()
    except errors.NotFound: pass
    
    try: 
        channel_id = ctx.channel.id
        embed = Embed(title=f"__**Канал {ctx.channel.name} добавлен в автоотправку расписания:**__", color=0x00ee00)
    except: 
        channel_id = ctx.author.id
        embed = Embed(title=f"__**Пользователь {ctx.author.name} добавлен в автоотправку расписания:**__", color=0x00ee00)
        
    await add_channel(channel_id, group, h, m)
    embed.add_field(name=f'**Время отправки:**', value=f'> {h}:{m:0>2}', inline=False)
    await ctx.send(ctx.user.mention, embed=embed, reference=ctx.message)

@wBot.slash_command(name="noautosend", description="Отключить автоматическую отправку расписания")
async def no_auto_send(ctx):
    try: await ctx.delete()
    except errors.NotFound: pass
    
    try: 
        channel_id = ctx.channel.id
        embed = Embed(title=f"__**Канал {ctx.channel.name} удалён из автоотправки расписания**__", color=0xee0000)
    except: 
        channel_id = ctx.author.id
        embed = Embed(title=f"__**Пользователь {ctx.author.name} удалён из автоотправки расписания**__", color=0xee0000)
    
    await delete_msg_channel(channel_id)
    
    embed.add_field(name=f'**Расписание больше не будет отправляться автоматически**', value=':(', inline=False)
    await ctx.send(ctx.user.mention, embed=embed, reference=ctx.message)

@loop(seconds=30)
async def bg_task():
    if not wBot.is_ready(): return
    
    try:
        channels = await check_send_time()
        for channel_id, group in channels:
            channel = wBot.get_channel(channel_id) or (await wBot.fetch_user(channel_id))
            
            if channel is None:
                print(f"Could not find channel with ID {channel_id}")
                await delete_msg_channel(channel_id); return
                
            now = datetime.datetime.now()
            wd = now.weekday()
            if now.hour > 12: wd = (wd+1)%7
            await couples(channel, group, (await get_days(channel))[wd])
            
    except FileNotFoundError as e: print(getattr(e, 'message', str(e)))
    except IndexError as e: print(getattr(e, 'message', str(e)))
    except Exception: traceback.print_exc()

bg_task.start()
event_loop = asyncio.get_event_loop()
event_loop.run_until_complete(wBot.start(token=os.environ.get('ds_token')))