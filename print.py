import os
import subprocess
import time

from configobj import ConfigObj

config = ConfigObj('config.cfg')

for filename in os.listdir(config['輸出資料夾']):
    if not filename.endswith('.pdf'):
        continue

    filepath = os.path.join(config['輸出資料夾'], filename)
    subprocess.call(r'C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe /s /h /t "{:s}"'.format(filepath))
    time.sleep(1.0)
    print(filepath)