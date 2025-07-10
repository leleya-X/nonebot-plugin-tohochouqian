
from pathlib import Path
from nonebot.adapters.onebot.v11 import Bot, GroupMessageEvent, Message, MessageSegment
from nonebot.params import CommandArg
from nonebot.plugin import PluginMetadata
from nonebot.plugin.on import on_command
import random
import asyncio
import openpyxl

PLUGIN_DIR = Path(__file__).parent


__plugin_meta__ = PluginMetadata(
    name="抽签插件",
    description="128张东方存幻神签抽取nonebotbot小程序，支持选择签编号。",
    usage="输入#chou或者#抽签可以抽一张东方幻存神签\n使用方式：\n输入#chou或者#抽签来进行抽签\n#chou 序号 可以显示相应序号的签\n高清晰度，附有简单中文说明)",
    type="application",
    homepage="https://github.com/leleya-X/nonebot-plugin-tohochouqian",
    supported_adapters={"~onebot.v11"},
    extra={
        "author": "YourName",
        "version": "0.1.0",
    }
)

chouqian = on_command("chou", aliases={"抽签", "chouqian"}, priority=5, block=True)

@chouqian.handle()
async def handle_chouqian(
    bot: Bot,
    event: GroupMessageEvent,
    args: Message = CommandArg()
):

    bid = args.extract_plain_text().strip()

    # 获取用户昵称
    user_name = event.sender.card or event.sender.nickname
    await asyncio.sleep(random.random())  # 随机延迟防并发

    # 处理签号选择
    if bid.isdigit() and 1 <= int(bid) <= 128:
        chouL = int(bid)
        text = f"{user_name}，您选择显示第{bid}号签"
    elif bid:
        chouL = random.randint(1, 128)
        text = f"{user_name}，只有128张哦，现在随便给你抽一张～"
    else:
        chouL = random.randint(1, 128)
        text = f"{user_name}，这是您抽到的签!"

    # 构建消息
    message = Message()


    message.append(MessageSegment.text(text))
    message.append(MessageSegment.image(file=f"file://{basepath}/chou/{chouL}.png"))


    wb = openpyxl.load_workbook(f"{basepath}/trs.xlsx")
    message.append(MessageSegment.text(wb.active[f"A{chouL}"].value))
    message.append(MessageSegment.text("\n"))
    message.append(MessageSegment.text(wb.active[f"B{chouL}"].value + r"|" + wb.active[f"C{chouL}"].value))
    message.append(MessageSegment.text("\n"))
    message.append(MessageSegment.text(wb.active[f"D{chouL}"].value))
    wb.close()
   ## await MessageFactory(message).send()
    await chouqian.finish(message)

