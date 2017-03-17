import os

from comtypes.client import CreateObject
from configobj import ConfigObj
from natsort import natsorted

WD_FORMAT_PDF = 17

config = ConfigObj('config.cfg')
word = CreateObject('Word.Application')

for path in natsorted(os.listdir(config['輸出資料夾'])):
    if not path.endswith('.docx'):
        continue

    print(path)

    (name, _) = os.path.splitext(path)

    in_path = os.path.join(
        config['輸出資料夾'],
        path,
    )
    out_path = os.path.join(
        config['輸出資料夾'],
        '{:s}.pdf'.format(name),
    )

    document = word.Documents.Open(in_path)
    document.SaveAs(out_path, FileFormat=WD_FORMAT_PDF)
    document.Close()

word.Quit()