#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
REPORT
This outputs a csv file with the hostname of the switch and today's date.
This script uses netmiko to gather information via ssh.
---Order of commands---
'show ip interface brief | ex OK'
'show interfaces description | exclude Protocol Description'
'show cdp nei' + ip_int_br[i][0]
'show int ' + ip_int_br[i][0] + ' capabilities | in Type|Duplex'
'show int ' + ip_int_br[i][0] + ' switchport | in Administrative Mode|Operational Mode|Access Mode VLAN'
'show mac address-table interface ' + ip_int_br[i][0] + ' | ex Vlan|-|Table|Total'
'show run | in hostname'
"""

__author__ = 'Ryan Murray'
__version__ = '1.0'
__maintainer__ = 'Ryan Murray'
__email__ = 'ryan.murray.570@gmail.com'
__contributors__ = 'Ryan Murray, Lakota Meagher'
__status__ = 'Prototype'

import re
import sys
import traceback
from datetime import date
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar

import openpyxl
import pandas as pd
from mac_vendor_lookup import MacLookup
from netmiko import ConnectHandler
from styleframe import StyleFrame

def show_error(self, *args):
    err = traceback.format_exception(*args)
    messagebox.showerror('Exception', err)

    IPLabel = Label(root, text='Enter IP Address of Device: ', padx=10, width=25, anchor=W)
    usernameLabel = Label(root, text='Enter Username: ', padx=10, width=25, anchor=W)
    passwordLabel = Label(root, text='Enter Password: ', padx=10, width=25, anchor=W)

    ipAddress = Entry(root)
    username = Entry(root)
    password = Entry(root, show='*')

    show = Label(root, text='                       ')
    show.grid(row=5, column=0)

    submitButton = Button(root, text='Submit', command=lambda: submit(ipAddress.get(), username.get(), password.get()))
    submitButton.grid(row=4, column=0)

    IPLabel.grid(row=0, column=0)
    usernameLabel.grid(row=1, column=0)
    passwordLabel.grid(row=2, column=0)

    ipAddress.grid(row=0, column=1)
    username.grid(row=1, column=1)
    password.grid(row=2, column=1)
    root.update()

def submit(ipAddress, username, password):
    global show
    global progress

    show.destroy()
    show = Label(root, text=f'Atempting to connect to {ipAddress}')
    show.grid(row=5, column=0)
    root.update()

    net_device = {
        'device_type': 'cisco_ios',
        'ip': ipAddress,
        'username': username,
        'password': password,
    }

    # Logs into the networking device
    net_connect = ConnectHandler(**net_device)
    show.destroy()
    show = Label(root, text=f'Connection to {ipAddress} successful!')
    show.grid(row=5, column=0)
    root.update()

    # Enters 'show ip int br' and puts it in ip_int_br
    show_ip_int_br = net_connect.send_command('show ip interface brief | ex OK')
    show_ip_int_br = show_ip_int_br.lstrip('\n')
    show_ip_int_br = show_ip_int_br.rstrip('\n')
    ip_int_br = [x.split() for x in show_ip_int_br.split('\n')]

    # Enters 'show int desc' and puts it in int_desc
    show_int_desc = net_connect.send_command('show interfaces description | exclude Protocol Description')
    int_desc = []
    for x in show_int_desc.split('\n'):
        int_desc.append(x[55:])

    # Defines the table to use and inputs headers
    table = []

    for i in range(len(ip_int_br)):
        show.destroy()
        show = Label(root, text=f'Processing port {str(i)} out of {str(len(ip_int_br))}')
        show.grid(row=5, column=0)
        MAX = int(len(ip_int_br))
        progress_var = DoubleVar()
        progress = Progressbar(root, orient=HORIZONTAL, length=200, variable=progress_var, maximum=MAX)
        progress_var.set(i)
        progress.grid(row=6, column=0)
        root.update()

        # If the interface is a vlan it will not try to find cdp nei, speed, duplex, switchport info, mac, and oui lookup
        if re.search('Vlan.+', ip_int_br[i][0]):

            # Adds gathered information to table without mac
            table.append([ip_int_br[i][0],                   # Interface
                '',                                          # Speed
                '',                                          # Duplex
                '',                                          # Switchport
                '',                                          # Vlan
                ip_int_br[i][1],                             # IP Address
                ip_int_br[i][4] + '/' + ip_int_br[i][5],     # Status
                '',                                          # Connected Mac
                '',                                          # OUI Lookup
                int_desc[i],                                 # Description
                ''                                           # CDP Neighbors
            ])
        else:
            
            # Gathers CDP information
            cdp_nei = net_connect.send_command(f'show cdp nei {ip_int_br[i][0]}')
            cdp_nei = cdp_nei.lstrip('\n')
            cdp_nei = cdp_nei[289:]
            cdp_nei = cdp_nei.split(' ')
            cdp_nei = cdp_nei[0].strip()

            # Gathers Speed and Duplex information 
            show_speed_duplex = net_connect.send_command(f'show int {ip_int_br[i][0]} capabilities | in Type|Duplex')
            show_speed_duplex = show_speed_duplex[1:]
            speed_duplex = [x.split() for x in show_speed_duplex.split('\n')]
            try:
                speed = speed_duplex[0][1]
                duplex = speed_duplex[1][1]
            except:
                speed = ''
                duplex = ''

            # Gathers Etherchannel and Trunk information
            show_switchport = net_connect.send_command(f'show int {ip_int_br[i][0]} switchport | in Administrative Mode|Operational Mode|Access Mode VLAN')
            show_switchport = show_switchport.lstrip('\n')
            show_switchport = show_switchport.rstrip('\n')
            switchport = show_switchport.split('\n')
            trunk_access = switchport[0]
            trunk_access = trunk_access[21:]
            if trunk_access == 'trunk' or trunk_access == ' trunk':
                trunk_access = switchport[1]
                trunk_access = trunk_access[18:]
            try:
                vlan = switchport[2]
                vlan = vlan[18:]
            except:
                vlan = 'not switchable'

            # Gathers mac address information
            mac_table = net_connect.send_command(f'show mac address-table interface {ip_int_br[i][0]} | ex Vlan|-|Table|Total')
            mac_table = mac_table.lstrip('\n')
            mac_table = mac_table.rstrip('\n')
            mac_table = mac_table.split('\n')
            for m in range(len(mac_table)):
                mac = mac_table[m].split()
                try:
                    mac = mac[1]
                    oui = MacLookup().lookup(str(mac))
                except:
                    mac = ''
                    oui = ''
                # Adds gathered information to table
                table.append([ip_int_br[i][0],                   # Interface
                    speed,                                       # Speed
                    duplex,                                      # Duplex
                    trunk_access,                                # Switchport
                    vlan,                                        # Vlan
                    ip_int_br[i][1],                             # IP Address
                    ip_int_br[i][4] + '/' + ip_int_br[i][5],     # Status
                    mac,                                         # Connected Mac
                    oui,                                         # OUI Lookup
                    int_desc[i],                                 # Description
                    cdp_nei                                      # CDP Neighbors
                ])

    # Gathers hostname from network device
    hostname = net_connect.send_command('show run | in hostname')
    hostname = hostname[9:]

    # Disconnects from device
    net_connect.disconnect()

    show.destroy()
    show = Label(root, text=f'Saving report as {hostname} {str(date.today())}.xlsx')
    show.grid(row=5, column=0)
    root.update()
    # Save table as a xlsx
    StyleFrame.A_FACTOR = 3
    columns = ['Interface','Speed','Duplex','Switchport','Vlan','IP Address','Status','Connected Mac','OUI Lookup','Description','CDP Neighbors']
    export_file_path = filedialog.asksaveasfilename(initialfile=hostname + ' ' + str(date.today()),defaultextension='.xlsx')
    df = pd.DataFrame(data=table, columns=columns)
    excel_writer = StyleFrame.ExcelWriter(export_file_path)
    sf = StyleFrame(df)
    sf.to_excel(
        excel_writer=excel_writer, 
        best_fit=columns,
        columns_and_rows_to_freeze='B2', 
        row_to_add_filters=0,
    )
    excel_writer.save()

    messagebox.showinfo(title='Finished!', message='Time to party!')

if __name__ == '__main__':
    root = Tk()
    root.title('Network Device Report')

    Tk.report_callback_exception = show_error

    IPLabel = Label(root, text='Enter IP Address of Device: ', padx=10, width=25, anchor=W)
    usernameLabel = Label(root, text='Enter Username: ', padx=10, width=25, anchor=W)
    passwordLabel = Label(root, text='Enter Password: ', padx=10, width=25, anchor=W)

    ipAddress = Entry(root)
    username = Entry(root)
    password = Entry(root, show='*')

    show = Label(root, text='                       ')
    show.grid(row=5, column=0)

    submitButton = Button(root, text='Submit', command=lambda: submit(ipAddress.get(), username.get(), password.get()))
    submitButton.grid(row=4, column=0)

    IPLabel.grid(row=0, column=0)
    usernameLabel.grid(row=1, column=0)
    passwordLabel.grid(row=2, column=0)

    ipAddress.grid(row=0, column=1)
    username.grid(row=1, column=1)
    password.grid(row=2, column=1)

    root.mainloop()
