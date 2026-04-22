from ping3 import ping
import socket
import psutil
import wmi
#Version: 0.1.18

#Windows 7 Check

# Класс с методами для тестирования состояния устройства
class Utils():
    #Конструтор
    def __init__(self):
        self.version = '0.1.15'

    #Данный закрытый метод нужнен, чтобы делать сразу несколько запросов пинг до узла.
    #Ведь если просто вызывать ping(), то это всего 1 ICMP запрос, а для уверености
    #надо хотя бы 4.
    def __CheckPing(self, node:str, timeout:int=4, count_ping:int=4):
        ping_list = []

        while count_ping > 0:
            count_ping -= 1
            p = ping(node, timeout=timeout)

            if str(p) == str('False') or str(p) == str('0.0'):
                ping_list.append(False)
            else:
                ping_list.append(True)

        count_ok = 0
        for result in ping_list:
            if result == True:
                count_ok += 1
        if count_ok >= 2:
            return True
        else:
            return False
    #Данный закрытый метод нужен для проверки доступности порта.
    #Возращает: в случаи доступности порта True, недоступности False
    #В случае, если не будет найден узел, будет возращено False, с
    #сообщением 'socket.gaierror'.
    def __CheckMailPort(self, node:str, port:int, timeout:int=4):
        set_soket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        set_soket.settimeout(timeout)

        try:
            set_soket.connect((node, port))
            return {'status': True, 'msg': ''}
        except socket.gaierror:
            return {'status': False, 'msg': 'socket.gaierror'}
        except:
            return {'status': False, 'msg': ''}
        finally:
            set_soket.close()

    # Данный закрытый метод нужен для проверки доступности корпоративного узла.
    # Возращает: в случаи доступности узла True, недоступности False
    def __CheckCorpNode(self, node:str, port:int=135, timeout:int=4):
        set_soket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        set_soket.settimeout(timeout)

        try:
            set_soket.connect((node, port))
            return {'status': True, 'msg': ''}
        except socket.gaierror:
            return {'status': False, 'msg': 'socket.gaierror'}
        except OSError as e:
            if str(e).split(']')[0] == '[WinError 10051':
                return {'status': False, 'msg': ''}
            else:
                return {'status': False, 'msg': ''}
        except Exception:
            return {'status': False, 'msg': ''}
        finally:
            set_soket.close()
    #В данном закрытом методе для определения типа устройства,
    #используется след логика. Если у стройства есть встроенная батерея,
    #то тогда это ноутбук. (При подвлкюченых ИБП по USB, для отслеживания их состояния.
    # Может такое быть, что метод определит неверно тип устройство, но это очень маловероятно)
    def __CheckNotebook(self):
        try:
            cwmi = wmi.WMI()
            battery = len(cwmi.Win32_Battery())
        except:
            return False

        if battery >= 1:
            return True
        else:
            return False

    #Метод для получения имени устройства
    def GetHostName(self):
        try:
            return {'status': True, 'place': 'hostname', 'data': str(socket.gethostname()), 'msg':''}
        except Exception as e:
            return {'status': False, 'place': 'hostname', 'data': '', 'msg':str(e)}
    #Метод для получения IP адреса сетевых адаптеров Ethernet и VPN
    def GetHostIP(self):
        win = wmi.WMI()

        def CneckOnVPN():
            for net in win.Win32_NetworkAdapterConfiguration(IPEnabled=True):
                if str(net.Description).lower() == 'Shrew Soft Virtual Adapter'.lower():
                    return {"VPN": str(net.IPAddress[0])}
            return False
        def CneckOnEthernet():
            for iface, ipv in psutil.net_if_addrs().items():
                for ip in ipv:
                    if ip.family == socket.AF_INET:
                        if str(iface).lower() == 'Ethernet'.lower():
                            return {"Ethernet": str(ip.address)}
            return False

        try:
            result_vpn = CneckOnVPN()
            result_ethernet = CneckOnEthernet()

            if result_vpn != False or result_ethernet != False:
                return {'status': True,
                        'place': 'ip',
                        'data': result_vpn if result_vpn != False else result_ethernet,
                        'msg': ''}
            else:
                return {'status': False,
                        'place': 'ip',
                        'data': {},
                        'msg': 'Сетевые адаптеры \nEthernet, VPN - не найдены.'}
        except Exception as e:
            return {'status': False, 'place': 'ip', 'data': {}, 'msg': str(e)}
    #Данный метод пингует надежные узлы сети для проверки доступа в интернет
    def CheckNet(self, timeout:int=4):
        try:
            netconn1 = self.__CheckPing('google.com', timeout=int(timeout))
            netconn2 = self.__CheckPing('yandex.ru', timeout=int(timeout))
        except OSError as e:
            if str(e).split(']')[0] == '[WinError 10051':
                return {'status': False,
                        'place': 'network',
                        'data': {'google.com': False, 'yandex.ru': False},
                        'msg': 'Ни один надёжный узел \nнедоступен, доступ в интернет отсутствует.'}
            else:
                return {'status': False,
                        'place': 'network',
                        'data': {'google.com': False, 'yandex.ru': False},
                        'msg': str(e)}
        except Exception as e:
            return {'status': False,
                    'place': 'network',
                    'data': {'google.com': False, 'yandex.ru': False},
                    'msg': str(e)}

        if netconn1 != False:
            if netconn2 != False:
                return {'status': True,
                        'place': 'network',
                        'data': {'google.com': True, 'yandex.ru': True},
                        'msg': ''}
            else:
                return {'status': True,
                        'place': 'network',
                        'data': {'google.com': True, 'yandex.ru': False},
                        'msg': 'Один надёжный узел недоступен'}
        else:
            if netconn2 != False:
                return {'status': True,
                        'place': 'network',
                        'data': {'google.com': False, 'yandex.ru': True},
                        'msg': 'Один надёжный узел недоступен'}
            else:
                return {'status': False,
                        'place': 'network',
                        'data': {'google.com': False, 'yandex.ru': False},
                        'msg': 'Ни один надёжный узел \nнедоступен, доступ в интернет отсутствует.'}
    # Данный метод пингует надежные узлы сети для проверки доступа в корпоративную сеть
    def CheckCorpNet(self, timeout: int = 4, trust_node1: str = 'HO-TSC', trust_node2: str = 'VO-TSC'):
        result_node1 = self.__CheckCorpNode(node=trust_node1, timeout=timeout)
        result_node2 = self.__CheckCorpNode(node=trust_node2, timeout=timeout)

        if result_node1['msg'] == '' or result_node2['msg'] == '':
            if result_node1['status'] == True or result_node2['status'] == True:
                return {'status': True,
                        'place': 'co_network',
                        'data': {'trust_node1': True, 'trust_node2': True},
                        'msg': ''}
            else:
                return {'status': False,
                        'place': 'co_network',
                        'data': {'trust_node1': False, 'trust_node2': False},
                        'msg': 'Ни один корпоративный \nузел недоступен'}
        else:
            return {'status': False,
                    'place': 'co_network',
                    'data': {'trust_node1': False, 'trust_node2': False},
                    'msg': 'Ни один корпоративный \nузел недоступен'}
    #Проверяет доступность почтового сервера и используемых им портов
    def CheckMail(self, timeout:int=4, mail_node:str='mxs.union-co.ru', imap_port:int=993, smtp_port:int=465):
        result_imap = self.__CheckMailPort(mail_node, imap_port, timeout)
        result_smtp = self.__CheckMailPort(mail_node, smtp_port, timeout)

        if result_imap['msg'] == '' and result_smtp['msg'] == '':
            if result_imap['status'] == True and result_smtp['status'] == True:
                return {'status': True,
                        'place': 'co_mail',
                        'data': {'mail_node': True, 'imap_port': result_imap['status'], 'smtp_port': result_smtp['status']},
                        'msg': ''}
            else:
                return {'status': False,
                        'place': 'co_mail',
                        'data': {'mail_node': True, 'imap_port': result_imap['status'], 'smtp_port': result_smtp['status']},
                        'msg': 'Один или несколько \nпортов недоступны.'}
        else:
            return {'status': False,
                    'place': 'co_mail',
                    'data': {'mail_node': False, 'imap_port': False, 'smtp_port': False},
                    'msg': 'Почтовый сервер недоступен.'}
    #Проверят подключенны ли базовые устройства ввода информации и в каком кол-во
    def CheckInputDevice(self):
        try:
            cwmi = wmi.WMI()
            keyboard = len(list(cwmi.Win32_Keyboard()))
            mouse = len(list(cwmi.Win32_PointingDevice()))
            notebook = self.__CheckNotebook()
        except Exception as e:
            return {'status': False,
                    'place': 'device',
                    'data': {'keyboard': 0, 'mouse': 0},
                    'msg': str(e)}

        if notebook == True:
            if int(keyboard) <= 2 and int(mouse) <= 2:
                return {'status': True,
                        'place': 'device',
                        'data': {'keyboard': keyboard, 'mouse': mouse},
                        'msg': ''}
            elif int(keyboard) == 0 and int(mouse) == 0:
                return {'status': False,
                        'place': 'device',
                        'data': {'keyboard': keyboard, 'mouse': mouse},
                        'msg': 'Устройств не обнаружено'}
            else:
                return {'status': False,
                        'place': 'device',
                        'data': {'keyboard': keyboard, 'mouse': mouse},
                        'msg': 'Одно из устройств не обнаружено \nили подключено лишнее устройство.'}
        else:
            if int(keyboard) == 1 and int(mouse) == 1:
                return {'status': True,
                        'place': 'device',
                        'data': {'keyboard': keyboard, 'mouse': mouse},
                        'msg': ''}
            elif int(keyboard) == 0 and int(mouse) == 0:
                return {'status': False,
                        'place': 'device',
                        'data': {'keyboard': keyboard, 'mouse': mouse},
                        'msg': 'Устройств не обнаружено'}
            else:
                return {'status': False,
                        'place': 'device',
                        'data': {'keyboard': keyboard, 'mouse': mouse},
                        'msg': 'Одно из устройств не обнаружено \nили подключено лишнее устройство.'}

#Проверка работы методов
# if __name__ == '__main__':
#     while True:
#         u = Utils()
#         print(f'Version: {u.version}', '\n')
#         print("HostName: ", u.GetHostName(), '\n')
#         print("HostIP: ", u.GetHostIP(), '\n')
#         print("Состояние сети: ", u.CheckNet(), '\n')
#         print("Состояние корп сети: ", u.CheckCorpNet(), '\n')
#         print('Состояние почты: ', u.CheckMail(), '\n')
#         print('Девайсы: ', u.CheckInputDevice(), '\n')
#
#         print('\n')
#         input()
