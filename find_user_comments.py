# -*- coding: utf-8 -*-
"""
Bilibili 视频评论爬取（带登录态，仅主评论，游标翻页无上限）
功能：运行后提示用户输入 cookie、BV 号、最大条数、请求延迟，
     通过游标接口抓取视频全量主评论（突破普通分页 5000 条上限），
     边抓边写入本地 SQLite 数据库（断点续传），
     完成后导出到 Excel 表格。

优化：
  - 每批评论实时写入 SQLite，崩溃/断网最多丢最后一批（约20条）
  - 重新运行时自动从上次中断处续抓
  - 写入 Excel 前过滤非法字符，避免 openpyxl 报错

依赖：
    pip install bilibili-api-python openpyxl httpx

打包成 exe（在 Windows 上执行）：
    pip install pyinstaller
    pyinstaller -F -n bilibili_comment_tool ^
      --collect-all bilibili_api ^
      --collect-all httpx ^
      --collect-all httpcore ^
      --collect-all anyio ^
      find_user_comments.py
"""

import asyncio
import httpx
import hashlib
import time
import re
import sqlite3
import os
from collections import defaultdict

from bilibili_api import video, Credential

# ============ 安全下限 ============
最小请求间隔秒 = 0.5
默认请求间隔秒 = 1.0
默认最大条数   = 0        # 0 = 不限制，抓到底
# =================================

MIXIN_KEY_ENC_TAB = [
    46,47,18,2,53,8,23,32,15,50,10,31,58,3,45,35,
    27,43,5,49,33,9,42,19,29,28,14,39,12,38,41,13,
    37,48,7,16,24,55,40,61,26,17,0,1,60,51,30,4,
    22,25,54,21,56,59,6,63,57,62,11,36,20,34,44,52,
]

# Excel / XML 禁止的控制字符（保留 \t \n \r）
_非法字符正则 = re.compile(
    r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f"
    r"\ud800-\udfff"        # 孤立代理对
    r"\ufffe\uffff]"
)

def 清理非法字符(text: str) -> str:
    """过滤 Excel/XML 不允许的字符，保留正常中文、emoji、换行"""
    if not isinstance(text, str):
        return str(text)
    return _非法字符正则.sub("", text)


# ── 交互输入 ────────────────────────────────────────────────────────

def 询问(提示文字, 默认值=None):
    后缀 = f"（默认 {默认值}）" if 默认值 is not None else ""
    答 = input(f"{提示文字}{后缀}：").strip()
    if not 答 and 默认值 is not None:
        return str(默认值)
    return 答


def 收集配置():
    print("=" * 50)
    print("   Bilibili 视频评论统计工具（全量版）")
    print("=" * 50)
    print("提示：cookie 从浏览器获取 —— 登录B站 → F12 →")
    print("      Application → Cookies → https://www.bilibili.com")
    print("      建议使用小号，cookie 等同登录凭证，请勿外传。")
    print("-" * 50)

    sessdata = 询问("请输入 SESSDATA")
    while not sessdata:
        print("  SESSDATA 不能为空，请重新输入。")
        sessdata = 询问("请输入 SESSDATA")

    bili_jct = 询问("请输入 bili_jct")
    buvid3   = 询问("请输入 buvid3")

    # 选择评论区类型
    print("-" * 50)
    print("请选择要抓取的评论区类型：")
    print("  1 = 视频评论区（输入 BV 号）")
    print("  2 = 动态评论区 - 转发/纯文字类（输入动态 ID）")
    print("  3 = 动态评论区 - 图文相册类（输入动态 ID）")
    print("  提示：动态 ID 是动态链接 t.bilibili.com/ 后面那串数字。")
    print("        若类型 2 抓不到评论，请改用类型 3，反之亦然。")
    while True:
        类型选择 = 询问("请选择（1 / 2 / 3）", "1")
        if 类型选择 in ("1", "2", "3"):
            break
        print("  请输入 1、2 或 3。")

    bvid = ""        # 仅视频用
    oid_直填 = None   # 动态用，直接就是 oid
    if 类型选择 == "1":
        评论类型 = 1
        bvid = 询问("请输入目标视频 BV 号（如 BV1fSGJ69EbE）")
        while not bvid:
            print("  BV 号不能为空，请重新输入。")
            bvid = 询问("请输入目标视频 BV 号")
        目标标识 = bvid
    else:
        评论类型 = 17 if 类型选择 == "2" else 11
        while True:
            动态id = 询问("请输入动态 ID（链接 t.bilibili.com/ 后的数字串）")
            动态id = 动态id.strip()
            if 动态id.isdigit():
                oid_直填 = int(动态id)
                break
            print("  动态 ID 应该是一串纯数字，请重新输入。")
        目标标识 = f"dyn{oid_直填}"

    while True:
        原始 = 询问("最多抓取多少条主评论（0 表示不限制，抓到底）", 默认最大条数)
        try:
            最大条数 = int(原始)
            break
        except ValueError:
            print("  请输入一个整数。")

    while True:
        原始 = 询问(f"请输入请求延迟秒数（不低于 {最小请求间隔秒}）", 默认请求间隔秒)
        try:
            延迟 = float(原始)
        except ValueError:
            print("  请输入一个数字。")
            continue
        if 延迟 < 最小请求间隔秒:
            print(f"  延迟过低，已自动提升到安全下限 {最小请求间隔秒} 秒。")
            延迟 = 最小请求间隔秒
        break

    # 输出模式
    print("-" * 50)
    print("输出模式：")
    print("  1 = 完整保留每条评论内容（默认）")
    print("  2 = 仅统计评论数量（不保留内容，表格更精简）")
    while True:
        模式 = 询问("请选择输出模式（1 或 2）", "1")
        if 模式 in ("1", "2"):
            保留内容 = (模式 == "1")
            break
        print("  请输入 1 或 2。")

    输出文件名 = 询问("请输入导出文件名（不含扩展名）", "B站评论统计结果")

    print("-" * 50)
    return {
        "凭证": Credential(sessdata=sessdata, bili_jct=bili_jct, buvid3=buvid3),
        "评论类型": 评论类型,        # 1=视频 17=转发动态 11=图文动态
        "bvid": bvid,               # 仅视频类型有值
        "oid_直填": oid_直填,        # 动态类型的 oid（视频为 None）
        "目标标识": 目标标识,        # 用于缓存文件命名
        "最大条数": 最大条数,
        "请求间隔秒": 延迟,
        "保留内容": 保留内容,
        "输出文件名": 输出文件名,
    }


# ── SQLite 落盘 ─────────────────────────────────────────────────────

def 初始化数据库(db路径: str):
    """建表（已存在则跳过），返回 connection"""
    conn = sqlite3.connect(db路径)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS 评论 (
            id       INTEGER PRIMARY KEY AUTOINCREMENT,
            uid      TEXT NOT NULL,
            昵称     TEXT,
            内容     TEXT,
            抓取时间 INTEGER
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS 进度 (
            键   TEXT PRIMARY KEY,
            值   TEXT
        )
    """)
    conn.commit()
    return conn


def 读取进度(conn) -> tuple[int, int]:
    """读取上次中断时的 next_offset 和已抓条数"""
    cur = conn.execute("SELECT 键, 值 FROM 进度")
    rows = {k: v for k, v in cur.fetchall()}
    return int(rows.get("next_offset", 0)), int(rows.get("总数", 0))


def 保存进度(conn, next_offset: int, 总数: int):
    conn.execute("INSERT OR REPLACE INTO 进度 VALUES ('next_offset', ?)", (str(next_offset),))
    conn.execute("INSERT OR REPLACE INTO 进度 VALUES ('总数', ?)",        (str(总数),))
    conn.commit()


def 批量写入评论(conn, 批次: list):
    """批次：[(uid, 昵称, 内容), ...]"""
    now = int(time.time())
    conn.executemany(
        "INSERT INTO 评论 (uid, 昵称, 内容, 抓取时间) VALUES (?, ?, ?, ?)",
        [(uid, 昵称, 内容, now) for uid, 昵称, 内容 in 批次]
    )
    conn.commit()


def 从数据库读取结果(conn) -> dict:
    """聚合成 {uid: {"昵称": str, "评论": [...]}}"""
    结果 = defaultdict(lambda: {"昵称": "", "评论": []})
    cur = conn.execute("SELECT uid, 昵称, 内容 FROM 评论 ORDER BY id")
    for uid, 昵称, 内容 in cur.fetchall():
        结果[uid]["昵称"] = 昵称 or ""
        结果[uid]["评论"].append(内容 or "")
    return 结果


# ── WBI 签名 ────────────────────────────────────────────────────────

def 获取混淆密钥(img_key: str, sub_key: str) -> str:
    raw = img_key + sub_key
    return "".join(raw[i] for i in MIXIN_KEY_ENC_TAB)[:32]


def wbi签名(params: dict, img_key: str, sub_key: str) -> dict:
    mixin_key = 获取混淆密钥(img_key, sub_key)
    params["wts"] = int(time.time())
    sorted_params = dict(sorted(params.items()))
    query = "&".join(
        f"{k}={re.sub(r'[!#$&+,/:;=?@\\[\\]]', '', str(v))}"
        for k, v in sorted_params.items()
    )
    params["w_rid"] = hashlib.md5((query + mixin_key).encode()).hexdigest()
    return params


async def 获取wbi密钥(凭证) -> tuple[str, str]:
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/120.0.0.0 Safari/537.36",
        "Referer": "https://www.bilibili.com",
    }
    cookies = {
        "SESSDATA": 凭证.sessdata,
        "bili_jct": 凭证.bili_jct,
        "buvid3":   凭证.buvid3,
    }
    async with httpx.AsyncClient(headers=headers, cookies=cookies) as client:
        resp = await client.get("https://api.bilibili.com/x/web-interface/nav")
        data = resp.json()
    wbi_img = data["data"]["wbi_img"]
    def _key(url):
        return url.rsplit("/", 1)[-1].split(".")[0]
    return _key(wbi_img["img_url"]), _key(wbi_img["sub_url"])


# ── 游标翻页抓取（边抓边存） ─────────────────────────────────────────

async def 抓取评论(cfg: dict, conn: sqlite3.Connection):
    凭证    = cfg["凭证"]
    评论类型 = cfg["评论类型"]
    bvid    = cfg["bvid"]
    最大条数 = cfg["最大条数"]
    间隔    = cfg["请求间隔秒"]

    # 确定 oid 与 referer
    if 评论类型 == 1:
        v = video.Video(bvid=bvid, credential=凭证)
        oid = await v.get_aid() if asyncio.iscoroutinefunction(v.get_aid) else v.get_aid()
        referer = f"https://www.bilibili.com/video/{bvid}"
        目标描述 = f"视频 {bvid}（aid={oid}）"
    else:
        oid = cfg["oid_直填"]
        referer = f"https://t.bilibili.com/{oid}"
        类型名 = "转发/纯文字动态" if 评论类型 == 17 else "图文相册动态"
        目标描述 = f"{类型名}（oid={oid}，type={评论类型}）"

    print("正在获取 WBI 签名密钥……")
    img_key, sub_key = await 获取wbi密钥(凭证)

    # 读取断点
    next_offset, 总数 = 读取进度(conn)
    if 总数 > 0:
        print(f"检测到上次进度：已抓 {总数} 条，从游标 {next_offset} 处续抓……")
    else:
        print(f"开始抓取 {目标描述} 的全量主评论……")

    限制说明 = f"最多 {最大条数} 条" if 最大条数 > 0 else "不限制，抓到底"
    print(f"抓取上限：{限制说明}")
    print("提示：随时按 Ctrl+C 可主动停止，进度自动保存，下次运行从断点续抓。\n")

    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/120.0.0.0 Safari/537.36",
        "Referer": referer,
    }
    cookies = {
        "SESSDATA": 凭证.sessdata,
        "bili_jct": 凭证.bili_jct,
        "buvid3":   凭证.buvid3,
    }

    轮次 = 1

    try:
        async with httpx.AsyncClient(headers=headers, cookies=cookies, timeout=15) as client:
            while True:
                if 最大条数 > 0 and 总数 >= 最大条数:
                    print(f"已达到条数上限（{最大条数} 条），停止抓取。")
                    break

                params = wbi签名(
                    {"oid": oid, "type": 评论类型, "mode": 3, "next": next_offset},
                    img_key, sub_key,
                )

                try:
                    resp = await client.get(
                        "https://api.bilibili.com/x/v2/reply/wbi/main",
                        params=params,
                    )
                    data = resp.json()
                except asyncio.CancelledError:
                    raise   # 交给外层统一处理
                except Exception as e:
                    print(f"第 {轮次} 次请求失败：{e}")
                    print("网络异常，进度已保存，重新运行可从此处续抓。")
                    break

                code = data.get("code", -1)
                if code != 0:
                    print(f"接口返回错误码 {code}：{data.get('message', '')}，停止抓取。")
                    print("进度已保存，重新运行可从此处续抓。")
                    break

                cursor_info = data["data"].get("cursor", {})
                replies     = data["data"].get("replies") or []

                if not replies:
                    print("已抓取全部主评论，结束。")
                    break

                # 整理本批并写入数据库
                本批 = []
                for r in replies:
                    if 最大条数 > 0 and 总数 + len(本批) >= 最大条数:
                        break
                    本批.append((
                        r["member"]["mid"],
                        清理非法字符(r["member"]["uname"]),
                        清理非法字符(r["content"]["message"]),
                    ))

                批量写入评论(conn, 本批)
                总数 += len(本批)

                next_offset = cursor_info.get("next", 0)
                is_end      = cursor_info.get("is_end", True)

                # 每批抓完立刻保存游标
                保存进度(conn, next_offset, 总数)

                print(f"第 {轮次:>4} 次请求完成，本批 {len(本批)} 条，累计 {总数} 条（已落盘）……")
                轮次 += 1

                if is_end or (最大条数 > 0 and 总数 >= 最大条数):
                    print("已到达评论区末尾，抓取完成。")
                    break

                # 用 asyncio.sleep 等待，Ctrl+C 会在此处被立即响应
                await asyncio.sleep(间隔)

    except (asyncio.CancelledError, KeyboardInterrupt):
        print(f"\n⚑ 已手动停止。本次共抓取 {总数} 条，进度已保存。")
        print("下次运行时选择续抓，将从断点继续。\n")
        return   # 直接返回，不打印"本次共抓取"

    print(f"\n本次共抓取 {总数} 条主评论。\n")


# ── 报表 & 导出 ──────────────────────────────────────────────────────

def 生成报表(结果, 保留内容=True):
    行列表 = []
    for uid, 信息 in 结果.items():
        评论 = 信息["评论"]
        if 保留内容:
            评论内容 = 清理非法字符(
                "\n".join(f"{i+1}. {c}" for i, c in enumerate(评论))
            )
            行列表.append([uid, 信息["昵称"], len(评论), 评论内容])
        else:
            行列表.append([uid, 信息["昵称"], len(评论)])
    行列表.sort(key=lambda x: x[2], reverse=True)
    return 行列表


def 导出xlsx(行列表, 文件名, 保留内容=True):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill

    print("正在导出 Excel，请稍候……")
    wb = Workbook()
    ws = wb.active
    ws.title = "评论统计"

    if 保留内容:
        表头 = ["用户UID", "用户昵称", "评论数量", "评论内容"]
        宽度 = [16, 20, 10, 70]
    else:
        表头 = ["用户UID", "用户昵称", "评论数量"]
        宽度 = [16, 20, 12]

    ws.append(表头)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for 行 in 行列表:
        # 逐格再过滤一次，彻底防止残留非法字符
        净行 = [清理非法字符(str(v)) if isinstance(v, str) else v for v in 行]
        ws.append(净行)

    for i, w in enumerate(宽度, start=1):
        ws.column_dimensions[chr(64 + i)].width = w
    for 行 in ws.iter_rows(min_row=2):
        for cell in 行:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    路径 = f"{文件名}.xlsx"
    wb.save(路径)
    print(f"已导出 Excel：{路径}")


# ── 入口 ─────────────────────────────────────────────────────────────

async def 异步主流程():
    cfg = 收集配置()
    目标标识 = cfg["目标标识"]

    # 数据库文件名与目标绑定（视频用BV号，动态用dyn+oid），续抓时自动对应
    db路径 = f"评论缓存_{目标标识}.db"
    db是新的 = not os.path.exists(db路径)
    conn = 初始化数据库(db路径)

    if not db是新的:
        _, 已有条数 = 读取进度(conn)
        if 已有条数 > 0:
            print(f"\n发现缓存文件 {db路径}，已存有 {已有条数} 条评论。")
            选择 = 询问("是否从上次中断处续抓？（y=续抓 / n=清空重新抓）", "y").lower()
            if 选择 == "n":
                conn.close()
                os.remove(db路径)
                conn = 初始化数据库(db路径)
                print("已清空缓存，重新开始抓取。\n")

    try:
        await 抓取评论(cfg, conn)
    except KeyboardInterrupt:
        pass   # 抓取函数内部已处理并打印提示
    except Exception as e:
        print(f"\n抓取过程出错：{e}")
        print("进度已保存，重新运行可从中断处续抓。")

    结果 = 从数据库读取结果(conn)
    conn.close()

    if not 结果:
        print("未抓取到任何评论，请检查 BV 号、cookie 或网络。")
        return

    # 无论是正常结束还是手动停止，都询问是否导出
    _, 当前条数 = 读取进度(sqlite3.connect(db路径))
    导出选择 = 询问(f"\n数据库中共有 {当前条数} 条评论，是否立即导出 Excel？（y=导出 / n=跳过保留缓存下次续抓）", "y").lower()
    if 导出选择 != "y":
        print(f"已跳过导出，缓存保留在：{db路径}")
        print("下次运行时选择续抓，完成后再导出。")
        return

    行列表 = 生成报表(结果, cfg["保留内容"])
    print("=" * 50)
    print(f"统计完成：共 {len(行列表)} 个用户。")
    模式说明 = "完整保留评论内容" if cfg["保留内容"] else "仅统计评论数量"
    print(f"输出模式：{模式说明}")
    print("=" * 50)
    导出xlsx(行列表, cfg["输出文件名"], cfg["保留内容"])

    # 导出成功后询问是否删除缓存
    删除 = 询问(f"\n是否删除缓存文件 {db路径}？（y=删除 / n=保留）", "y").lower()
    if 删除 == "y":
        os.remove(db路径)
        print("缓存已删除。")
    else:
        print(f"缓存已保留在：{db路径}")


def main():
    try:
        asyncio.run(异步主流程())
    except Exception as e:
        print(f"\n程序出错：{e}")
    finally:
        input("\n按回车键退出……")


if __name__ == "__main__":
    main()
