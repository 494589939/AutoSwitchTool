#!/usr/bin/python3
# -*- coding: utf-8 -*-
import sys
from time import sleep

import paramiko
import xlwings as xw
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog

from Ui_h3c import Ui_Form  # 调用生成的.py文件


# 通过python paramiko编写一个能够备份H3C交换机的脚本，可以使用备份前的配置文件，自动备份到tftp服务器上
# 模板带有的参数为ip，端口，用户，旧密码，新密码，tftp地址
# 运行流程main_UI-->change_passwd-->openfile-->len_rows-->ssh_h3c


# 0.备份命令
def backup():
    cmd = 'backup startup-configuration to {tftp_ip} {ip}.cfg \n'
    switch_cmd(cmd)


# 1.查看透明端口
def mac_authentication():
    ip_info = ''
    cmd = """
        screen-length dis \n
        dis cu int \n
        """
    # 清空ui.textBrowser文本内容
    ui.textBrowser.clear()
    # 调用openfile函数获取交换机信息与行数
    switch_lists, total_rows = openfile()

    # 对交换机信息与行数进行判断如果为空就判断为告警
    if switch_lists is None:
        ui.textBrowser.setText(str("文件未上传或上传识别失败请检查EXCEL！"))
    else:
        # 根据获取的行数减去开通一行生成进度条最大值
        ui.progressBar.setRange(0, total_rows - 1)
        # 打印需要执行的数量
        ui.textBrowser.append("合计交换机数量:" + str(total_rows - 1))
        for ip, port, user, old_pwd, new_pwd, su_pwd in switch_lists:
            try:
                # new_cmd = cmd.format(tftp_ip=None, ip=ip, port=port, user=None, old_pwd=None, new_pwd=None)
                outputs = ssh_h3c(ip, port, user, old_pwd, new_pwd, cmd, su_pwd)
                # 显示文本清洗数据
                for i in str(outputs).split("#"):
                    if 'mac-authentication' in i:
                        for s in i.split('\n'):
                            if 'GigabitEthernet' in s:
                                ip_info = ip_info + ip + ' ' + s + '\n'
                ui.textBrowser.append(ip_info)
                outcome = ("++" + ip + '执行成功！')
            except Exception as e:
                # 打印结果并跳过此次执行
                print(e)
                outcome = ("--" + ip + '执行失败:' + str(e))
                continue
            finally:
                # 不论结果如何都更新进度条
                ui.progressBar.setValue(ui.progressBar.value() + 1)
                # 将结果打印至UI
                ui.textBrowser.append(str(outcome))
                # 刷新UI编辑框
                qt_app.processEvents()


# 0.改密码命令
def change_passwd():
    # su切换用户时pwd一定要注意字符串的空格
    cmd = """
su\n
{su_pwd}\n
system-view\n
local-user {user} class manage\n
password simple {new_pwd}\n
quit\n
save force\n
"""
    switch_cmd(cmd)


# 1.调用Ui打开文件管理器
def openfile():
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        # 调用QT5的QFileDialog获取文件路径
        file_path, _ = QFileDialog.getOpenFileName(
            None, '选择文件', '.', 'Excel files(*.xlsx , *.xls)')
        wb = app.books.open(file_path)
        # 打开表一
        sht = wb.sheets[0]
        total_rows = len_rows(wb, sht)
        # 如果实际数据只要一行数就将列表再包一层
        if total_rows == 2:
            rng = [sht.range('A{0}:F{1}'.format(2, total_rows)).value]
        else:
            rng = sht.range('A{0}:F{1}'.format(2, total_rows)).value
        wb.save()
        # 关闭
        wb.close()
        # 杀死进程
        app.kill()
        ui.progressBar.setValue(0)
        return rng, total_rows
    except Exception as e:
        print(e)
        rng, total_rows = None, None
        return rng, total_rows


# 2.获取excel总行数
def len_rows(wb, sheet_index):
    sheet = wb.sheets[sheet_index]
    rng = sheet.range('A1').expand('table')
    nrows = rng.rows.count
    return nrows


# 3.交换机配置总命令
def switch_cmd(cmd):
    # 清空ui.textBrowser文本内容
    ui.textBrowser.clear()
    # 获取tftp地址需要先打开tftp64软件
    tftp_ip = ui.lineEdit.text()
    # 调用openfile函数获取交换机信息与行数
    switch_lists, total_rows = openfile()

    # 对交换机信息与行数进行判断如果为空就判断为告警
    if switch_lists is None:
        ui.textBrowser.setText(str("文件未上传或上传识别失败请检查EXCEL！"))
    else:
        # 根据获取的行数减去开通一行生成进度条最大值
        ui.progressBar.setRange(0, total_rows - 1)
        # 打印需要执行的数量
        ui.textBrowser.append("合计交换机数量:" + str(total_rows - 1))
        # print(switch_lists)
        for ip, port, user, old_pwd, new_pwd, su_pwd in switch_lists:
            try:
                new_cmd = cmd.format(tftp_ip=tftp_ip, ip=ip, port=port, user=user, old_pwd=old_pwd, new_pwd=new_pwd,
                                     su_pwd=su_pwd)
                ssh_h3c(ip, port, user, old_pwd, new_pwd, new_cmd, su_pwd)
                outcome = ("++" + ip + '执行成功！')
            except Exception as e:
                # 打印结果并跳过此次执行
                print(e)
                outcome = ("--" + ip + '执行失败:' + str(e))
                continue
            finally:
                # 不论结果如何都更新进度条
                ui.progressBar.setValue(ui.progressBar.value() + 1)
                # 将结果打印至UI
                ui.textBrowser.append(str(outcome))
                # 刷新UI编辑框
                qt_app.processEvents()


# 4.发送交换机命令并连接ssh执行
def ssh_h3c(ip, port, user, old_pwd, new_pwd, new_cmd, su_pwd):
    # 创建ssh连接
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(ip, int(port), user, old_pwd)
    ui.textBrowser.append('<------------------成功连接上:' + ip + '------------------>')

    # 发送命令
    command = ssh.invoke_shell()
    command.send(new_cmd)
    # 必须设置等待时间
    sleep(2)

    # # 解析为UTF-8编码
    output = command.recv(65535).decode('UTF-8')
    # # 显示文本
    # ui.textBrowser.setText(str(output))

    # 关闭连接
    ssh.close()
    return output


if __name__ == '__main__':
    qt_app = QApplication(sys.argv)
    MainWindow = QWidget()
    ui = Ui_Form()
    ui.setupUi(MainWindow)
    MainWindow.show()
    # 开始制作ui的接口
    ui.pushButton.clicked.connect(backup)
    ui.pushButton_1.clicked.connect(change_passwd)
    ui.pushButton_2.clicked.connect(mac_authentication)
    sys.exit(qt_app.exec_())
