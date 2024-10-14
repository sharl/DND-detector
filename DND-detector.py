# -*- coding: utf-8 -*-
import time
import io
import binascii
import threading

import schedule
from pystray import Icon, Menu, MenuItem
from PIL import Image
from win32gui import EnumWindows, GetClassName
from SwitchBot import SwitchBot

NAME = 'Do Not Disturb detector'
WNDCLASSNAMES = [
    'SQEX.CDev.Engine.Framework.MainWindow',
    'Qt663QWindowIcon',         # OBS Studio
]
deviceID = '806599370E22'


class DQX:
    def __init__(self):
        self.running = False
        self.existing = False
        self.paused = False
        self.switchbot = SwitchBot()
        # API access count
        self.count = 0

        self.on_image = Image.open(io.BytesIO(binascii.unhexlify(ON.replace('\n', '').strip())))
        self.off_image = Image.open(io.BytesIO(binascii.unhexlify(OFF.replace('\n', '').strip())))
        menu = Menu(
            MenuItem('Pause', self.pauseApp, checked=lambda _: self.paused),
            Menu.SEPARATOR,
            MenuItem('Exit', self.stopApp),
        )
        self.app = Icon(name=NAME, title=NAME, icon=self.off_image, menu=menu)
        self.detect()

    def detect(self):
        if self.paused:
            print('paused')
            return

        self.proc = None

        def EnumWindowsProc(hWnd):
            for WNDCLASSNAME in WNDCLASSNAMES:
                if GetClassName(hWnd) == WNDCLASSNAME:
                    self.proc = hWnd
                    break

        EnumWindows(lambda hWnd, _: EnumWindowsProc(hWnd), None)

        if self.proc is None:
            if self.existing:
                self.switchbot.set_device_power(deviceID, 'off')
                self.count = self.count + 1
                self.existing = False
                self.app.icon = self.off_image
                print('not running, turn off plug', self.count)
        else:
            if not self.existing:
                self.switchbot.set_device_power(deviceID, 'on')
                self.count = self.count + 1
                self.existing = True
                self.app.icon = self.on_image
                print('running, turn on plug', self.count)
        self.app.update_menu()

    def pauseApp(self):
        self.paused = not self.paused

    def runSchedule(self):
        # SwitchBot API rate limit: 10000 times per day
        # https://github.com/OpenWonderLabs/SwitchBotAPI/blob/main/README.md#request-limit
        # 86400 / 60 -> 1440
        schedule.every(60).seconds.do(self.detect)

        while self.running:
            schedule.run_pending()
            time.sleep(1)

    def stopApp(self):
        self.running = False
        self.app.stop()

    def runApp(self):
        self.running = True

        task_thread = threading.Thread(target=self.runSchedule)
        task_thread.start()

        self.app.run()


ON = """
89504e470d0a1a0a0000000d4948445200000010000000100803000000282d0f530000000467414d410000b18f0bfc6105000000206348524d00007a
26000080840000fa00000080e8000075300000ea6000003a98000017709cba513c00000183504c5445ffff00fefe00fdfd01f3f211eaee26e5f024e5
f025e7f025e9f025e8ef24eaef24ecef24e9f126e9ee24e6f025e4ee23e5f023eaee25f4f310a4a74a3720356d1e276b1e29581f29461e274e202947
252e35242c491a2445252d601e2970283369222a362136adb048999a3c6a1220f5bcb7e0787ad2474bbb4d4ecb5b5b8d2223791819c56667a32729de
6d6ee99393ed9895671624a2a33c9e8b2ead7080fc8c8ef19499fac7c8fcdbdafbc2c4d72f3ae58083ffdddbec686df5a6aaf7c1c0ffd0cc832538a2
a0389c9133984152ffd2ccf8a7a7f39294f7afaefdded9e92a36f7a5a8f48b8ff9a8aaf6999af5a3a4ffa7a5802134a2a1389ca04536000a96383880
23257c101379101288191d73090e87191c7d03049d2227a120258c181b8217154e0e1fa3a33fb7b74152556c554c5755515e53546052535f54505d54
515d575b66584e5b574e5a58546055545e515067bfbf3ef7f706dddc20dbdd23dbdc22dcdc22dbdb21dadc21dbda20dbdb22dcdb21dadc22dbdc23dd
dd20f9f904ffffffc2ad91bd00000001624b47448065bd9e680000000774494d4507e80a0d130011fac9cee8000000a34944415418d3636020081899
500023a60a66165636760e4e2e6e1e5e3e7e01412106611151317109492969195939790545250665155535750d4d2d6d1d5d3d7d0343230663135333
730b4b2b6b1b5b3b7b0747270667175737770f4f2f6f1f5f3fff80c02086e090d0b0f088c8a8e898d8b8f884c42486e494d4b4f48cccf4acec9cdcbc
fc824286a2e292d2b2f28acaaaea9a9ad2daba7a0654a7313212e174740000eaab2155766a23650000002574455874646174653a6372656174650032
3032342d31302d31335431303a30303a32322b30393a3030ae8909ad0000002574455874646174653a6d6f6469667900323032342d31302d31335431
303a30303a31372b30393a3030036399550000000049454e44ae426082
"""

OFF = """
89504e470d0a1a0a0000000d4948445200000010000000100803000000282d0f53000000206348524d00007a26000080840000fa00000080e8000075
300000ea6000003a98000017709cba513c00000132504c54450000008d8d8d8484848181828282838383838484848282838282828181818080808282
828484858b8b8b6e6e6e3030302929292828282727272828282a2a2b29292a3232327070705e5e5e6363645d5d5d6464646060606564656363636767
677878785f5f605c5c5c5c5c5d5b5b5b5b5b5c5b5b5c5a5a5b5b5b5b5c5c5d6161627a7a7a7777787474757373747272737070717070717070717171
72717172727273747475777778151517282b2f1f222514161815171917181b08090a09090a1a1c1f0d0e0f1b1d2026292d23262a1516172527292427
2b2a2d312e3236363a3f2f32370f101221232733373c1e2023292c3032363b2f333818191a1c1e202e313632353b0e0f1125282c2325291818191414
141313151011110d0d0e0c0c0d1010110a0a0a0909090e0f100c0d0d000000c8a034a00000003874524e53002c46464646454545464546472aa6fbf8
f8f8f9f8f8fb9baca1aba2aca2aca18bf8f9f8f8f9f8f9f9f9f680154246444546474546464112bb504e6b00000001624b47440088051d4800000007
74494d4507e80a0d12342f81404a830000006a744558745261772070726f66696c6520747970652061707031000a617070310a20202020202033340a
343934393261303030383030303030303031303033313031303230303037303030303030316130303030303030303030303030303437366636663637
36633635303030300aa75f8a99000000954944415418d36360a0026064626661656363e76065e5e4e2e6e165e0e31710141414101412161614111513
6790b0b0b4b2b6b1b5b37770747276719564907273f7f0f4f2f6f1f5f30f080c0a9666900909f5700f080b8f887476b6748f9265908b8e898d8b4f48
744c4a48484e89926750505452565151555357d150d5d0d4d266d0d1d5d3373034323634343135313533a78657002b5518d20f4db87b000000257445
5874646174653a63726561746500323032342d31302d31335430393a35323a35332b30393a30306d039ce10000002574455874646174653a6d6f6469
667900323032342d31302d31335430393a35323a34372b30393a303024bb00d00000000049454e44ae426082
"""


if __name__ == '__main__':
    DQX().runApp()
