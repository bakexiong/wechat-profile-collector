"""Excel → 企业微信 → 昵称及头像。"""

import argparse
import logging
from pathlib import Path
import sys

from workbook import collect, create_template


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("excel", type=Path, nargs="?", default=Path("phone_numbers.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("output/results.xlsx"))
    parser.add_argument("--pid", type=int, help="多个企业微信窗口时指定进程 ID")
    parser.add_argument("--interval", type=float, default=3, help="查询间隔秒数，默认 3")
    parser.add_argument("--init", action="store_true", help="创建空白输入模板并退出")
    args = parser.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    logging.getLogger("airtest").setLevel(logging.WARNING)
    try:
        if args.init:
            create_template(args.excel)
            logging.info("模板已创建：%s", args.excel)
            return 0
        if args.interval < 0:
            parser.error("--interval 不能小于 0")
        if sys.platform != "win32":
            raise RuntimeError("采集功能仅支持 Windows")
        from wecom import WeCom

        client = WeCom(pid=args.pid)
        collect(args.excel, args.output, client.lookup, interval=args.interval)
        return 0
    except KeyboardInterrupt:
        logging.info("已停止，已完成的记录保存在结果文件中")
        return 130
    except Exception as exc:
        logging.error("%s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
