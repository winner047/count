import os
import shutil

# 确保必要的目录存在
os.makedirs('/tmp/vercel_pandas', exist_ok=True)

# 复制必要的文件到临时目录（解决Vercel的文件系统限制）
if os.path.exists('app.py'):
    shutil.copy('app.py', '/tmp/')

print("Build completed successfully")