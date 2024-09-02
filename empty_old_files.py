# Empty recycle bin of files older than specified num weeks
import winshell, subprocess, sys, os, shutil, traceback, stat, errno
from datetime import datetime, timedelta, timezone


def handleRemoveReadonly(func, path, exc):
    if func in (os.rmdir, os.remove, os.unlink) and exc.errno == errno.EACCES:
        os.chmod(path, stat.S_IRWXU | stat.S_IRWXG | stat.S_IRWXO)  # 0777
        func(path)
    else:
        raise


subprocess.check_call([sys.executable, "-m", "pip", "install", "winshell"])
subprocess.check_call([sys.executable, "-m", "pip", "install", "pypiwin32"])

# deleting files older than this number of weeks
# must be negative
NUM_WEEKS = -52

# access the recycle bin
recycle_bin = winshell.recycle_bin()

# print the number of items in recycle bin
print(len(list(recycle_bin)), "items present in recycle bin")

# go through items, deleting old ones
for item in recycle_bin:
    if item.recycle_date() < (datetime.now(timezone.utc) + timedelta(weeks=NUM_WEEKS)):
        try:
            name = item.filename()
            if os.path.isdir(name):
                shutil.rmtree(name, onexc=handleRemoveReadonly)
            else:
                os.remove(name)
        except Exception as e:
            traceback.print_exception(e)

print(len(list(winshell.recycle_bin())), "items left in recycle bin")
