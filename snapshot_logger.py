from loguru import logger

logger.add("M:/CPP-Data/CBO Westbury Managers/LEADERSHIP/Bot Folder/Archive/Daily Snapshot Logs/{time:YYYY-MM-DD}.log",
           format="{time:YYYY-MM-DD at HH:mm:ss} | {level} - {message}",
           colorize=True, backtrace=True, diagnose=True, level='DEBUG', retention='90 days')