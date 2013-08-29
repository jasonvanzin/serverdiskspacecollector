__author__ = 'Jason Vanzin'

import sys
from PySide import QtCore, QtGui
from sdscgui import Ui_dlgMain
import win32ui
import pickle
import os
import wmi
import gspread
from itertools import count
import xml.etree.ElementTree as ET
import hashlib
from Crypto.Cipher import AES
import random
import math
import re
import win32com.client
import codecs
import webbrowser


serverlist = [] # global variable to store list of servers.
configpath = os.path.dirname(os.path.realpath(sys.argv[0])) + '/' #sets path to program folder


class mainwindow(Ui_dlgMain):
    exportformat = ''
    def __init__(self):
        super(mainwindow, self).__init__()
        self.setupUi(window)


    def preload(self):
        # checks to see is saved list exits. If it does, loads it. Also, checks to see if Excel is installed.
        # If not, it disables Excel option.
        try:
            serverlist = pickle.load(open(configpath + 'servers.dat', 'rb'))
            for server in serverlist:
                ui.lstServers.addItem(server)
        except:
            pass
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Quit()
        except:
            self.radioExcel.setEnabled(False)

    def buttonClicked(self):
        # Executes with any button click. Determines which button and executes code for that button.
        sender = self.sender().text()
        print(sender, 'clicked')
        if sender == 'Save As':
            # Opens a dialog box to specify where to save the file after determining if it should be an excel or csv file.
            if self.radioExcel.isChecked():
                filedlg = win32ui.CreateFileDialog( 1, ".xlsx", None, 0, "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*|")
                filedlg.DoModal()
                print(filedlg.GetPathName())
                self.txtFilename.setText(filedlg.GetPathName())
            elif self.radioCSV.isChecked():
                filedlg = win32ui.CreateFileDialog( 1, ".csv", None, 0, "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*|")
                filedlg.DoModal()
                print(filedlg.GetPathName())
                self.txtFilename.setText(filedlg.GetPathName())
            else:
                self.statusUpdate("Please choose and export format first.")

        if sender == 'Add Servers':
            # Adds servers from text box to the list box.
            if self.txtAddserver.displayText():
                self.statusUpdate('')
                self.lstServers.addItem(self.txtAddserver.displayText())
                self.txtAddserver.setText('')
            else:
                self.statusUpdate('You must enter a servername.')

        if sender == 'Save List':
            # Saves the list of servers to a file to be automatically loaded next time the program runs.
            serverlist.clear()
            counter = 0
            while counter < self.lstServers.count():
                serverlist.append(self.lstServers.item(counter).text())
                counter += 1
            serverfile = configpath + 'servers.dat'

            try:
                pickle.dump(serverlist, open(serverfile, 'wb'))
                self.statusUpdate('Server list saved!')
            except:
                self.statusUpdate('Server list failed to save. Check permissions on program directory.')

        if sender == "Save Credentials":
            # Saves the login information to an encrypted file, so the user can reload it next time the program runs.
            if not self.txtPassphrase.displayText():
                self.statusUpdate('Credentials could not be saved. Please enter a passphrase.')
                self.txtPassphrase.setFocus()
            else:
                try:
                    #configure XML elements
                    top = ET.Element('config')
                    server_username = ET.SubElement(top, 'server_username')
                    server_domain = ET.SubElement(top, 'server_domain')
                    server_passwd = ET.SubElement(top, 'server_passwd')
                    g_username = ET.SubElement(top, 'g_username')
                    g_passwd = ET.SubElement(top, 'g_passwd')
                    g_spreadsheet = ET.SubElement(top, 'g_spreadsheet')

                    # Asks user for input to popluate and XML elements
                    upassword, gpassword, passphrase = ui.getpasswords()
                    server_username.text = self.txtUsername.displayText()
                    server_domain.text = self.txtDomain.displayText()
                    server_passwd.text = upassword
                    g_username.text = self.txtGUsername.displayText()
                    g_passwd.text = gpassword
                    g_spreadsheet.text = self.txtGSpreadsheet.displayText()
                    password = passphrase



                    #configure encryption
                    mode = AES.MODE_CBC
                    key = hashlib.md5(password.encode('utf-16')).digest()
                    iv = bytes([random.randint(0,0xFF) for i in range(16)])
                    encryptor = AES.new(key, mode, iv)

                    #Encrypte and save data to file
                    stringtowrite = ET.tostring(top).decode("utf-8") #generate string to encrypt
                    stringlen = len(stringtowrite)
                    padding = (math.ceil(stringlen/16)*16) - stringlen #figure out padding
                    stringtowrite = stringtowrite + 'X' * padding #add padding
                    datatowrite = [encryptor.encrypt(stringtowrite), iv] #create list to write to file
                    pickle.dump(datatowrite, open(configpath + 'config.dat', 'wb'))
                    self.statusUpdate('Credentials have been encrypted and saved.')
                    self.txtPassphrase.clear()
                except:
                    self.statusUpdate('Credentials could not be saved. Check the permission on the program directory.')

        if sender == 'Load Credentials':
            # Loads the saved credentials from the encrypted file.
            if not self.txtPassphrase.displayText():
                self.statusUpdate('Please enter a passphrase used when credentials were saved.')
                self.txtPassphrase.setFocus()
            else:
                upassword, gpassword, passphrase = ui.getpasswords()
                configdict = get_config(passphrase)

                self.txtUsername.setText(configdict['server_username'])
                self.txtPassword.setText(configdict['server_passwd'])
                self.txtDomain.setText(configdict['server_domain'])
                self.txtGUsername.setText(configdict['g_username'])
                self.txtGPassword.setText(configdict['g_passwd'])
                self.txtGSpreadsheet.setText(configdict['g_spreadsheet'])
                self.txtPassphrase.clear()


        if sender == 'Scan':
            # Executes the scan of the servers listed in the server list box if checks pass.
            fieldsneeded = self.checkfields()
            self.statusUpdate('')
            if len(fieldsneeded) > 0:
                self.statusUpdate('Missing ' + str(fieldsneeded))
                print(str(fieldsneeded))
            elif self.lstServers.count() == 0:
                self.statusUpdate('No servers have been added to be scanned.')
            else:
                serverlist.clear()
                counter = 0
                while counter < self.lstServers.count():
                    serverlist.append(self.lstServers.item(counter).text())
                    counter += 1
                password = self.getpasswords()
                server_info_dict = scanservers(serverlist, self.txtUsername.displayText(), password[0], self.txtDomain.displayText())

                if self.exportformat == 'Excel':
                    self.statusUpdate('Saving to Excel....')
                    savetoexcel(server_info_dict, self.txtFilename.displayText())
                elif self.exportformat == 'CSV':
                    self.statusUpdate('Saving to CSV file ...')
                    savetocsv(server_info_dict, self.txtFilename.displayText())
                else:
                    clear_spreadsheet(self.txtGUsername.displayText(), password[1], self.txtGSpreadsheet.displayText())
                    self.statusUpdate('Saving to Google spreadsheet.....')
                    savetogoogle(server_info_dict, self.txtGUsername.displayText(), password[1], self.txtGSpreadsheet.displayText())

        if sender == 'Close':
            # exits program
            sys.exit(0)

    def checkfields(self):
        # Checks to see which fields have been left blank.
        fieldsneeded = []
        if not self.txtUsername.displayText(): fieldsneeded.append('Server Username')
        if not self.txtPassword.displayText(): fieldsneeded.append('Server Password')
        if not self.txtDomain.displayText(): fieldsneeded.append('Server Domain')
        if len(fieldsneeded) > 0:
            return fieldsneeded
        if not self.exportformat: fieldsneeded.append('Export Format')
        if len(fieldsneeded) > 0:
            return fieldsneeded
        if self.exportformat == 'google':
            if not self.txtGUsername.displayText(): fieldsneeded.append('Google Username')
            if not self.txtGPassword.displayText(): fieldsneeded.append('Google Password')
            if not self.txtGSpreadsheet.displayText(): fieldsneeded.append('Google Spreadsheet')
        else:
            if not self.txtFilename.displayText(): fieldsneeded.append('Filename')
        return fieldsneeded

    def getpasswords(self):
        # Gathers login information fields
        self.txtPassword.setEchoMode(QtGui.QLineEdit.Normal)
        self.txtGPassword.setEchoMode(QtGui.QLineEdit.Normal)
        self.txtPassphrase.setEchoMode(QtGui.QLineEdit.Normal)
        upassword = self.txtPassword.displayText()
        gpassword = self.txtGPassword.displayText()
        passphrase = self.txtPassphrase.displayText()
        self.txtPassword.setEchoMode(QtGui.QLineEdit.Password)
        self.txtGPassword.setEchoMode(QtGui.QLineEdit.Password)
        self.txtPassphrase.setEchoMode(QtGui.QLineEdit.Password)
        return upassword, gpassword, passphrase

    def openWebsite(self):
        # Opens website if url is clicked.
        webbrowser.open('http://essentialink.com')

    def removeServer(self):
        # removes server from list if server is double clicked
        itemno = self.lstServers.currentRow()
        self.lstServers.takeItem(itemno)


    def exportbtnUpdate(self, message):
        # Updates the text of the Save As button
        self.btnSaveas.setText(message)

    def radioChange(self):
        # Enables and disables various objects based on which radio button is chosen.
        if self.radioExcel.isChecked():
            self.exportbtnUpdate('Save As')
            self.exportformat = 'Excel'
            self.txtFilename.setEnabled(True)
            self.txtGUsername.setEnabled(False)
            self.txtGPassword.setEnabled(False)
            self.txtGSpreadsheet.setEnabled(False)
        elif self.radioCSV.isChecked():
            self.exportformat = 'CSV'
            self.exportbtnUpdate('Save As')
            self.txtFilename.setEnabled(True)
            self.txtGUsername.setEnabled(False)
            self.txtGPassword.setEnabled(False)
            self.txtGSpreadsheet.setEnabled(False)
        elif self.radioGoogle.isChecked():
            self.exportformat = 'google'
            self.exportbtnUpdate('Disabled')
            self.txtFilename.setEnabled(False)
            self.txtGUsername.setEnabled(True)
            self.txtGPassword.setEnabled(True)
            self.txtGSpreadsheet.setEnabled(True)

    def statusUpdate(self, resultmsg):
        # Updates the status message with whatever message is provided to the function.
        self.statusLabel.setText(resultmsg)


def get_config(password):
    # Open file and get encrypted data
    try:
        datafromfile = pickle.load(open(configpath + 'config.dat', 'rb'))
    except:
        err = 'Saved credentials not found.'
        return err

    # Configure encryption
    mode = AES.MODE_CBC

    key = hashlib.md5(password.encode('utf-16')).digest()
    decryptor = AES.new(key, mode, datafromfile[1])

    # Decrypt data from file.
    plaintext = decryptor.decrypt(datafromfile[0])

    # Pull out XML string only.
    try:
        rawdata = re.findall(r'(<config>.+</config>)', str(plaintext))[0]
    except:
        err = 'Unable to load saved credentials.'
        return err

    # Rebuild XML element tree.
    configdictionary = {}
    configdata = ET.fromstring(rawdata)

    # Build dictionary of configuration information.
    for child in configdata:
        configdictionary[child.tag] = child.text

    return configdictionary

def scanservers(serverlist, username, password, domain):
    # Scans the servers and builds a dictionary with the hard drive storage information.
    disk_info_dict = {}
    server_diskinfo_dict = {}
    print(serverlist)
    ui.progressBar.setEnabled(True)

    ui.statusUpdate('Scanning the servers.....')
    progresscounter = len(serverlist)
    for server in serverlist:
        try:
            c = wmi.WMI(computer=server, user=domain + '\\' + username, password=password, find_classes=False)
            for disk in c.Win32_LogicalDisk(DriveType=3):
                disk_info_dict[disk.Caption] = [round(int(disk.FreeSpace)/1024/1024/1024, 2),
                    round(int(disk.Size)/1024/1024/1024,2)]
            server_diskinfo_dict[server] = disk_info_dict
        except:
            server_diskinfo_dict[server] = "Could not communicate with server."
        ui.progressBar.setProperty("value", round(80 / progresscounter, 0))
        progresscounter = progresscounter - 1

    return server_diskinfo_dict

def savetoexcel(server_info_dict, filename):
    # Saves the information from the server_info_dict dictionary to the filename provided to Excel.
    ui.statusUpdate('Saving to excel....')
    constants = win32com.client.constants
    row_counter = count()
    next(row_counter)
    next(row_counter)
    excel = win32com.client.Dispatch("Excel.Application")
    book = excel.Workbooks.Add()
    book.SaveAs(filename)
    sheet = excel.Worksheets(1)
    sheet.Range('A1:E1').Value = ['Server', 'Drive', 'Total Space', 'Used Space', 'Free Space']

    for server in server_info_dict:
        if server_info_dict[server] == 'Could not communicate with server.':
            print(server, server_info_dict[server])
            row = next(row_counter)
            cellrange = 'A' + str(row) + ":B" + str(row)
            print('5', cellrange)
            sheet.Range(cellrange).Value = [server, server_info_dict[server]]
        else:
            for drive in server_info_dict[server]:
                print(server, drive)
                row = next(row_counter)
                cellrange = 'A' + str(row) + ":E" + str(row)
                sheet.Range(cellrange).Value = [server, drive, server_info_dict[server][drive][1],
                            round(server_info_dict[server][drive][1] - server_info_dict[server][drive][0], 2),
                            server_info_dict[server][drive][0]]



    book.Save()
    excel.Visible = True
    ui.statusUpdate("Server disk space have been saved to Excel.")
    ui.progressBar.setProperty("value", 100)


def savetocsv(server_info_dict, filename):
    # Saves the information from the server_info_dict dictionary to the filename provided to a csv file.
    ui.statusUpdate('Saving to csv....')
    csvfile = codecs.open (filename, "w", "latin-1")
    try:
        csvfile.write('Server,Drive,Total Space,Used Space, Free Space\n')
        for server in server_info_dict:
            if server_info_dict[server] == 'Could not communicate with server.':
                print(server, 'if section')
                stringtowrite = server + ',' + server_info_dict[server] + '\n'
                csvfile.write(stringtowrite)
            else:
                for drive in server_info_dict[server]:
                    print(server, drive, 'else section')
                    totalspace = server_info_dict[server][drive][1]
                    usedspace = round(server_info_dict[server][drive][1] - server_info_dict[server][drive][0], 2)
                    freespace = server_info_dict[server][drive][0]
                    stringtowrite = server + ',' + drive + ',' + str(totalspace)  + ',' + str(usedspace) + ',' + \
                                    str(freespace) + '\n'
                    print(stringtowrite)
                    csvfile.write(stringtowrite)
    except:
        ui.statusUpdate('Saving to csv failed.')
    csvfile.close()
    ui.statusUpdate('Server disk space saved to CSV file.')
    ui.progressBar.setProperty("value", 100)


def savetogoogle(server_info_dict, g_username, g_passwd, g_spreadsheet):
    # Saves the information from the server_info_dict dictionary to the filename provided to a Google spreadsheet.
    ui.statusUpdate("Saving servers to spreadsheet " + g_spreadsheet)
    print('1')
    gc = gspread.login(g_username, g_passwd)
    print('2')
    googlefile = gc.open(g_spreadsheet)

    print('3')
    worksheet = googlefile.sheet1
    print('4')
    row_counter = count()
    next(row_counter)
    next(row_counter)
    print(server_info_dict)
    for server in server_info_dict:
        if server_info_dict[server] == 'Could not communicate with server.':
            print(server, server_info_dict[server])
            row = next(row_counter)
            cellrange = 'A' + str(row) + ":B" + str(row)
            print('5', cellrange)
            cell_list = worksheet.range(cellrange)
            print('6')
            cell_list[0].value = server
            print('7')
            cell_list[1].value = server_info_dict[server]
            print('8')
            worksheet.update_cells(cell_list)
            print('9')
        else:
            for drive in server_info_dict[server]:
                print(server, drive)
                row = next(row_counter)
                cellrange = 'A' + str(row) + ":E" + str(row)
                print(cellrange)
                cell_list = worksheet.range(cellrange)
                print('10')
                cell_list[0].value = server
                cell_list[1].value = drive
                cell_list[2].value = server_info_dict[server][drive][1]
                cell_list[3].value = round(server_info_dict[server][drive][1] - server_info_dict[server][drive][0], 2)
                cell_list[4].value = server_info_dict[server][drive][0]
                worksheet.update_cells(cell_list)
    ui.statusUpdate("Servers have been saved to spreadsheet " + g_spreadsheet)
    ui.progressBar.setProperty("value", 100)


def clear_spreadsheet(g_username, g_password, g_spreadsheet):
    # Clears the spreadsheet if it exist or gives error if not. Spreadsheet must be created first.
    ui.statusUpdate("Clearing cells from spreadsheet " + g_spreadsheet)
    gc = gspread.login(g_username, g_password)
    try:
        googlefile = gc.open(g_spreadsheet)
    except:
        ui.statusUpdate('Spreadsheet doesn\'t exist. Please create blank spreadsheet first.')
        raise
    worksheet = googlefile.sheet1
    number_of_rows = worksheet.row_count
    cellrange = 'A1' + ":E" + str(number_of_rows)
    cell_list = worksheet.range(cellrange)
    for cell in cell_list:
        cell.value = ''
    worksheet.update_cells(cell_list)
    cell_list = worksheet.range('A1:E1')
    cell_list[0].value = 'Server'
    cell_list[1].value = 'Drive'
    cell_list[2].value = "Total Space"
    cell_list[3].value = 'Used Space'
    cell_list[4].value = 'Free Space'
    worksheet.update_cells(cell_list)
    ui.statusUpdate('')

if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    window = QtGui.QDialog()
    ui = mainwindow()
    ui.setupUi(window)
    ui.preload()
    window.show()

    sys.exit(app.exec_())

