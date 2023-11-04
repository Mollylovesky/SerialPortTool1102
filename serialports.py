import binascii

import serial
import serial.tools.list_ports
from time import sleep
import logging
from PyQt5.QtWidgets import QApplication
from mainwindowyuanban import *

"""
串口通信类
# 以16进制字节发送串口指令，以字符串形式接收返回值,返回值内需要有剩余字节数
# 实例化-调用autoconnect-checkrecv(send_data）
"""
logging.basicConfig(filename='./log.txt',
                    format='%(asctime)s-%(name)s-%(levelname)s-%(message)s-%(funcName)s:%(lineno)d',
                    level=logging.INFO)


class SerialPorts:
    def __init__(self, com=None, baudrate=9600, bytesize=8, parity='N', stopbits=1, xonxoff=False):
        self.ports_list = self.ports()
        self.baudrate = baudrate
        self.bytesize = bytesize
        self.parity = parity
        self.stopbits = stopbits
        self.xonxoff = xonxoff
        self.com = com
        self.checkhead = '5506'  # 指令协议规定的头


    def update(self, initdata=None, checkhead=None):
        if not initdata:
            self.initdata = initdata
        if not checkhead:
            self.checkhead = checkhead

    # 获取所有端口
    def ports(self):
        ports = []
        portslist = list(serial.tools.list_ports.comports())
        for port in portslist:
            ports.append(list(port)[0])
        return ports

    # 打开端口
    def opencom(self, com, baudrate=9600, bytesize=8, parity='N', stopbits=1, xonxoff=False):
        ser_open = serial.Serial(com, baudrate, bytesize, parity, stopbits, xonxoff)
        if ser_open.isOpen():
            logging.info(f"串口{com}打开成功")
            self.serial = ser_open
        else:
            logging.info(f"串口{com}打开失败")

    # 关闭端口，清空输入输出缓存
    def closecom(self):
        if self.serial.isOpen():
            self.clearcom()
            self.serial.close()
        if self.serial.isOpen():
            logging.info(f"串口{self.com}关闭失败")
        else:
            logging.info(f"串口{self.com}关闭成功")

    def clearcom(self):
        if self.serial.isOpen():
            self.serial.flushInput()
            self.serial.flushOutput()
            logging.info(f"串口{self.com}清空缓存成功")
        else:
            logging.info(f"串口{self.com}未打开")

    # 发送串口指令
    def send(self, send_data=None):
        if not self.serial.isOpen():
            logging.info("串口未连接！")
        if not send_data:
            send_data = self.initdata
            send_data_hex = self.initdata
            self.serial.write(send_data_hex.encode('utf-8'))
        else:
            send_data_hex = bytes.fromhex(send_data)
            self.serial.write(send_data_hex)
        logging.info(f"发送指令为：{send_data}")

    # 接收指令
    def recv(self, send_data, recvnum=1):
        data = []
        datas = self.serial.readline()
        hex_data = binascii.hexlify(datas).decode('utf-8')
        logging.info(f"接收数据{hex_data}")
        sleep(0.5)
        if hex_data == '':
            return 0
        else:
            for j in range(len(hex_data)):
                if hex_data[j: j + 4] == self.checkhead and datas[j+4:j+6] != send_data[4:6]:  # 查找返回值中与头一致的字符串
                    data.append(hex_data[j:j + 12])
                    logging.info(f"解析返回值{hex_data[j:j + 12]}")
            if len(data) == recvnum:
                return data
            elif len(data) == 2:
                recvnum = recvnum + 1
                return data, recvnum
            else:
                return 0, 0

    def datacut(self, data):
        datacuts = []
        datacuts_int = []
        i = 0
        while i < len(data) - 1: #len:字符串的长度
            datacuts.append(data[i: i + 2])
            datacuts_int.append(int(data[i: i + 2], 16))#把十六进制转换成了十进制
            i += 2
        return datacuts, datacuts_int

    def datacut1(self, data):
        datacuts = []
        i = 0
        while i < len(data) - 1: #len:字符串的长度
        # while i < 12 - 1:  # len:字符串的长度
            datacuts.append(data[i: i + 2])
            i += 2
        return datacuts
