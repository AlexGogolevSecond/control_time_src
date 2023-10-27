# pip install win11toast ; для win10 видимо win10toast
# пример: https://thepythoncorner.com/posts/2018-08-01-how-to-create-a-windows-service-in-python/
import os
from datetime import datetime
import socket
import time
from win11toast import toast
import win32serviceutil
import servicemanager
import win32event
import win32service
import sqlite3
import getpass
import subprocess
from typing import Optional


TIMEOUT = 10

class SMWinservice(win32serviceutil.ServiceFramework):
    """Base class to create winservice in Python"""
    _svc_name_ = 'control_time'
    _svc_display_name_ = 'Control time'
    _svc_description_ = 'Control time'
    @classmethod
    def parse_command_line(cls):
        """ClassMethod to parse the command line"""
        win32serviceutil.HandleCommandLine(cls)

    def __init__(self, args):
        """Constructor of the winservice"""
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)

    def SvcStop(self):
        """Called when the service is asked to stop"""
        self.stop()
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        """Called when the service is asked to start"""
        self.start()
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        self.main()

    def start(self):
        """
        Override to add logic before the start
        eg. running condition
        """
        pass

    def stop(self):
        """
        Override to add logic before the stop
        eg. invalidating running condition
        """
        pass

    def main(self):
        """Main class to be ovverridden to add logic"""
        pass

class PythonCornerExample(SMWinservice):

    def start(self):
        self.isrunning = True

    def stop(self):
        self.isrunning = False

    def main(self):
        while self.isrunning:
            self.control_time()
            time.sleep(50)

    def get_current_user_os(self) -> Optional[str]:
        """Определяем текущего пользователя ОС (Windows)
        Returns:
            Optional[str]: имя пользователя
        """
        res = subprocess.run(['powershell.exe', '$env:UserName'], 
                             capture_output=True)
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open("C:\ProgramData\control_time_srv\exception.txt",
                      'a', encoding='utf-8') as file:
            file.write(f'\n{now}: res.stdout: {res.stdout}')
        return res.stdout

    def control_time(self):
        """Проверяем текущего пользователя и текущее время"""
        try:
            with sqlite3.connect("C:\ProgramData\control_time_srv\log.db") as conn:
                # current_user = os.getlogin()
                # get_pass_get_user = getpass.getuser()
                current_user = self.get_current_user_os()
                if not current_user:
                    return
                
                hours = (0, 1, 2, 3, 4, 5, 6, 23, 24)
                target_time_delta : bool = datetime.now().hour in hours

                str_log = f'''current_user: {current_user}
                              target_time_delta: {target_time_delta}
                              datetime.now().hour: {datetime.now().hour}'''
                now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                cur = conn.execute(f"""insert into messages (datetime, message)
                                    values (\'{now}\', \'{str_log}\')""")
                cur.close()
                if 'dasha' in current_user.lower() and target_time_delta:
                    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    cur = conn.execute(f"""insert into messages (datetime, message)
                                    values (\'{now}\', \'вошли в условие по Даше, сейчас д.б. уведомление и выкл\')""")
                    cur.close()
                    toast(f'The computer will turn off in {TIMEOUT} seconds')
                    time.sleep(TIMEOUT)
                    os.system('shutdown -s')
        except Exception as ex:
            with open("C:\ProgramData\control_time_srv\exception.txt",
                      'a', encoding='utf-8') as file:
                file.write(str(ex))

if __name__ == '__main__':
    PythonCornerExample.parse_command_line()

'''
****************************************************

C:test> python PythonCornerExample.py install
Installing service PythonCornerExample
Service installed

In the future, if you want to change the code of your service, just modify it and reinstall the service with:
C:test> python PythonCornerExample.py update
Changing service configuration
Service updated


'''