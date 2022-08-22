#!/usr/bin/env python3

#
#  smfinspect.py
#
#  Created by yu2924 on 2016-12-05
#

import os, sys, struct, csv, traceback
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

class Error(Exception):
    pass

class SMFError(Error):
    def __init__(self, msg):
        self.message = msg

ccnametable = []
def resolveccname(cc):
    global ccnametable
    if len(ccnametable) == 0:
        ccnametable = [""] * 128
        path = os.path.abspath(os.path.join(os.path.dirname(sys.argv[0]), "Control Functions.txt"))
        with open(path, "rt", newline="", encoding="utf-8") as f:
            reader = csv.reader(f, delimiter="\t")
            for row in reader:
                index = int(row[0], 16)
                name = row[2]
                ccnametable[index] = name
    return ccnametable[cc] if (0 <= cc) and (cc < 128) else "#{0:02x}".format(cc)

mfidtable = []
def resolvemfid(ba):
    global mfidtable
    if len(mfidtable) == 0:
        path = os.path.abspath(os.path.join(os.path.dirname(sys.argv[0]), "Manufacturer ID Numbers.txt"))
        with open(path, "rt", newline="", encoding="utf-8") as f:
            reader = csv.reader(f, delimiter="\t")
            for row in reader:
                sig = bytes([int(b, 16) for b in row[0].split(" ")])
                name = row[2]
                mfidtable.append((sig, name))
    for e in mfidtable:
        sig = e[0]
        name = e[1]
        if ba[:len(sig)] == sig:
            return (name, len(sig))
    return ("", 0)

def resolvekeyname(sf, mi):
    KMAJ = [ "Cb", "Gb", "Db", "Ab", "Eb", "Bb", "F", "C", "G", "D", "A", "E", "B", "F#", "C#" ]
    KMIN = [ "Ab", "Eb", "Bb", "F", "C", "G", "D", "A", "E", "B", "F#", "C#", "G#", "D#", "A#" ]
    pk = KMIN if mi else KMAJ
    ssf = pk[sf + 7] if (-7 <= sf) and (sf <= 7) else "{0}".format(sf)
    smi = "min" if mi else "maj"
    return ssf + " " + smi;

def reads1(file):
    return struct.unpack("b", file.read(1))[0]

def readu1(file):
    return struct.unpack("B", file.read(1))[0]

def reads2(file):
    return struct.unpack(">h", file.read(2))[0]

def readu2(file):
    return struct.unpack(">H", file.read(2))[0]

def reads3(file):
    return (struct.unpack("b", file.read(1))[0] << 16) + (struct.unpack("B", file.read(1))[0] << 8) + struct.unpack("B", file.read(1))[0]

def readu3(file):
    return (struct.unpack("B", file.read(1))[0] << 16) + (struct.unpack("B", file.read(1))[0] << 8) + struct.unpack("B", file.read(1))[0]

def reads4(file):
    return struct.unpack(">l", file.read(4))[0]

def readu4(file):
    return struct.unpack(">L", file.read(4))[0]

def readvlq(file):
    v = 0
    u = 0x80
    while u & 0x80:
        u = struct.unpack("B", file.read(1))[0]
        v = (v << 7) + (u & 0x7f)
    return v;

def settablerow(table, row, deltatime, status, channel, data):
    itdt = QTableWidgetItem("{0}".format(deltatime))
    itdt.setToolTip(itdt.text())
    itdt.setTextAlignment(Qt.AlignRight|Qt.AlignVCenter)
    table.setItem(row, 0, itdt)
    its = QTableWidgetItem(status)
    its.setToolTip(its.text())
    its.setTextAlignment(Qt.AlignLeft|Qt.AlignVCenter)
    table.setItem(row, 1, its)
    itc = QTableWidgetItem(channel)
    itc.setToolTip(itc.text())
    itc.setTextAlignment(Qt.AlignRight|Qt.AlignVCenter)
    table.setItem(row, 2, itc)
    itd = QTableWidgetItem(data)
    itd.setToolTip(itd.text())
    itd.setTextAlignment(Qt.AlignLeft|Qt.AlignVCenter)
    table.setItem(row, 3, itd)

class MainWindow(QWidget):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.sysexes = []
        # widgets
        fm = self.fontMetrics()
        self.pathlabel = QLabel("")
        self.clearbtn = QPushButton("x", self)
        self.clearbtn.setFixedWidth(fm.height() * 2)
        self.clearbtn.clicked.connect(self.clearsmf)
        self.browsebtn = QPushButton("...", self)
        self.browsebtn.setFixedWidth(fm.height() * 2)
        self.browsebtn.clicked.connect(self.browsesmf)
        self.headerlabel = QLabel("")
        self.exportsyxbtn = QPushButton("export sysex", self)
        self.exportsyxbtn.setFixedWidth(fm.height() * 8)
        self.exportsyxbtn.clicked.connect(self.exportsyx)
        self.tabwidget = QTabWidget()
        # path sublayout
        pathlayout = QHBoxLayout()
        pathlayout.setSpacing(0)
        pathlayout.addWidget(self.pathlabel, 3)
        pathlayout.addWidget(self.clearbtn)
        pathlayout.addWidget(self.browsebtn)
        # header sublayout
        headerlayout = QHBoxLayout()
        headerlayout.setSpacing(0)
        headerlayout.addWidget(self.headerlabel, 3)
        headerlayout.addWidget(self.exportsyxbtn)
        # vertical stack
        vlayout = QVBoxLayout()
        vlayout.addLayout(pathlayout)
        vlayout.addLayout(headerlayout)
        vlayout.addWidget(self.tabwidget)
        # startup
        self.setLayout(vlayout)
        self.setGeometry(200, 200, 600, 800)
        self.setWindowTitle("SMF inspector")
        self.setAcceptDrops(True)
        self.show()
    def dragEnterEvent(self, event):
        mimedata = event.mimeData()
        if mimedata.hasUrls() and len(mimedata.urls())==1 and os.path.splitext(mimedata.urls()[0].toLocalFile())[1].lower()==".mid":
            event.accept()
        else:
            event.ignore()
    def dropEvent(self, event):
        try:
            mimedata = event.mimeData()
            path = mimedata.urls()[0].toLocalFile()
            self.pathlabel.setText(path)
            self.reload()
        except Exception as e:
            QMessageBox.warning(self, "dropEvent", traceback.format_exc(), QMessageBox.Ok)
    def browsesmf(self):
        try:
            dir = os.path.dirname(self.pathlabel.text())
            paths = QFileDialog.getOpenFileName(self, "Open MIDI", directory=dir, filter="*.mid")
            if 0 < len(paths):
                self.pathlabel.setText(paths[0])
                self.reload()
        except Exception as e:
            QMessageBox.warning(self, "browsesmf", traceback.format_exc(), QMessageBox.Ok)
    def clearsmf(self):
        self.pathlabel.setText("")
        self.headerlabel.setText("")
        self.tabwidget.clear()
        self.sysexes = []
    def reload(self):
        path = self.pathlabel.text()
        if path != "":
            self.loadsmf(path)
    def loadsmf(self, path):
        self.headerlabel.setText("")
        self.tabwidget.clear()
        self.sysexes = []
        try:
            with open(path, "rb") as f:
                # header
                hdrckid = f.read(4)
                if hdrckid != bytes(b"MThd"):
                    raise SMFError("MThd missing: " + str(hdrckid))
                hdrlen = readu4(f)
                if hdrlen != 6:
                    raise SMFError("MThd invalid size: {0}".format(hdrlen))
                smffmt = readu2(f) # format 0, 1, 2
                numtrks = readu2(f)
                division = reads2(f) # division of a quarter-note or {hibyte:frames per second, lobyte:division of a frame} pair
                hdrdesc = ""
                hdrdesc += "format={0} ".format(smffmt)
                hdrdesc += "numtracks={0} ".format(numtrks)
                if 0 <= division:
                    hdrdesc += "division={0}/qnote".format(division)
                else:
                    hdrdesc += "division={1}/frame, frames={0}/sec".format((-division >> 8) & 0xff, division & 0xff)
                self.headerlabel.setText(hdrdesc)
                # tracks
                for itrk in range(numtrks):
                    trkckid = f.read(4)
                    if trkckid != bytes(b"MTrk"):
                        raise SMFError("MTrk missing: " + str(trkckid))
                    trklen = readu4(f)
                    table = QTableWidget(parent=self)
                    self.tabwidget.addTab(table, "{0}".format(itrk))
                    table.setColumnCount(4)
                    table.setColumnWidth(0, 50)
                    table.setColumnWidth(1, 100)
                    table.setColumnWidth(2, 50)
                    table.setColumnWidth(3, 300)
                    table.setHorizontalHeaderLabels(("deltatime", "status", "channel", "data"));
                    table.setRowCount(1024)
                    # events
                    iev = 0
                    rs = 0
                    trkend = f.tell() + trklen
                    while f.tell() < trkend:
                        if table.rowCount() <= iev:
                            table.setRowCount(table.rowCount() * 2)
                        table.setRowHeight(iev, 20);
                        dt = readvlq(f)
                        bp = readu1(f)
                        f.seek(f.tell() - 1)
                        s = readu1(f) if bp & 0x80 else rs
                        if s < 0xf0:
                            sc = s & 0xf0
                            ch = s & 0x0f
                            rs = s
                            if sc == 0x80:
                                settablerow(table, iev, dt, "noteoff", "{0}".format(ch), "k={0} v={1}".format(readu1(f), readu1(f)))
                            elif sc == 0x90:
                                settablerow(table, iev, dt, "noteon", "{0}".format(ch), "k={0} v={1}".format(readu1(f), readu1(f)))
                            elif sc == 0xa0:
                                settablerow(table, iev, dt, "keypress", "{0}".format(ch), "k={0} v={1}".format(readu1(f), readu1(f)))
                            elif sc == 0xb0:
                                b = readu1(f)
                                settablerow(table, iev, dt, "control", "{0}".format(ch), "{0}({1}) v={2}".format(resolveccname(b), b, readu1(f)))
                            elif sc == 0xc0:
                                settablerow(table, iev, dt, "program", "{0}".format(ch), "v={0}".format(readu1(f)))
                            elif sc == 0xd0:
                                settablerow(table, iev, dt, "chpress", "{0}".format(ch), "v={0}".format(readu1(f)))
                            elif sc == 0xe0:
                                bl = readu1(f)
                                bh = readu1(f)
                                settablerow(table, iev, dt, "pitch", "{0}".format(ch), "v={0}".format(int((bh << 7) + bl) - 8192))
                            else:
                                settablerow(table, iev, dt, "invalid({0:02x})".format(s), "")
                                # f.seek(trkend)
                        elif s == 0xf0:
                            syxlen = readvlq(f)
                            syxdata = f.read(syxlen)
                            mfid,lmfid = resolvemfid(syxdata)
                            syxlen = syxlen + 1
                            syxdata = b"\xf0" + syxdata
                            syxelp = "..." if 128 < syxlen else ""
                            if 0 < lmfid:
                                settablerow(table, iev, dt, "sysex", "", "\"{1}\" length={0}: ".format(syxlen, mfid) + "".join(["{0:02x} ".format(b) for b in syxdata[:128]]) + syxelp)
                            else:
                                settablerow(table, iev, dt, "sysex", "", "length={0}: ".format(syxlen) + "".join(["{0:02x} ".format(b) for b in syxdata[:128]]) + syxelp)
                            self.sysexes.append(syxdata)
                        elif s == 0xf7:
                            syxlen = readvlq(f)
                            syxdata = f.read(syxlen)
                            syxelp = "..." if 128 < syxlen else ""
                            settablerow(table, iev, dt, "sysex.part", "", "length={0}: ".format(syxlen) + "".join(["{0:02x} ".format(b) for b in syxdata[:128]]) + syxelp)
                            self.sysexes.append(syxdata)
                        elif s == 0xff:
                            metatype = readu1(f);
                            metalen = readvlq(f);
                            metaelp = "..." if 128 < metalen else ""
                            if metatype == 0x00:
                                settablerow(table, iev, dt, "meta", "", "seqnum: " + "{0}".format(readu2(f)))
                            elif metatype == 0x01:
                                settablerow(table, iev, dt, "meta", "", "text: " + f.read(metalen)[:128].decode("utf-8", "ignore") + metaelp)
                            elif metatype == 0x02:
                                settablerow(table, iev, dt, "meta", "", "copyright: " + f.read(metalen)[:128].decode("utf-8", "ignore") + metaelp)
                            elif metatype == 0x03:
                                settablerow(table, iev, dt, "meta", "", "trackname: " + f.read(metalen)[:128].decode("utf-8", "ignore") + metaelp)
                            elif metatype == 0x04:
                                settablerow(table, iev, dt, "meta", "", "instname: " + f.read(metalen)[:128].decode("utf-8", "ignore") + metaelp)
                            elif metatype == 0x05:
                                settablerow(table, iev, dt, "meta", "", "lyric: " + f.read(metalen)[:128].decode("utf-8", "ignore") + metaelp)
                            elif metatype == 0x06:
                                settablerow(table, iev, dt, "meta", "", "marker: " + f.read(metalen)[:128].decode("utf-8", "ignore") + metaelp)
                            elif metatype == 0x07:
                                settablerow(table, iev, dt, "meta", "", "cuepoint: " + f.read(metalen)[:128].decode("utf-8", "ignore") + metaelp)
                            elif metatype == 0x08:
                                settablerow(table, iev, dt, "meta", "", "program: " + f.read(metalen)[:128].decode("utf-8", "ignore") + metaelp)
                            elif metatype == 0x09:
                                settablerow(table, iev, dt, "meta", "", "device: " + f.read(metalen)[:128].decode("utf-8", "ignore") + metaelp)
                            elif metatype == 0x20:
                                settablerow(table, iev, dt, "meta", "", "chprefix: " + "{0}".format(readu1(f)))
                            elif metatype == 0x21:
                                settablerow(table, iev, dt, "meta", "", "port: " + "{0}".format(readu1(f)))
                            elif metatype == 0x2f:
                                settablerow(table, iev, dt, "meta", "", "endoftrack")
                                f.seek(trkend)
                            elif metatype == 0x51:
                                tempo = readu3(f)
                                settablerow(table, iev, dt, "meta", "", "tempo: " + "bps={0:.2f} ({1} us/qn)".format(60 * 1000000 / tempo, tempo))
                            elif metatype == 0x54:
                                hr = readu1(f)
                                mn = readu1(f)
                                se = readu1(f)
                                fr = readu1(f)
                                ff = readu1(f)
                                settablerow(table, iev, dt, "meta", "", "smpte: " + "{0:02d}:{1:02d}:{2:02d}.{3:02d}:{4:02d}".format(hr, mn, se, fr, ff))
                            elif metatype == 0x58:
                                nn = readu1(f)
                                dd = readu1(f)
                                cc = readu1(f)
                                bb = readu1(f)
                                settablerow(table, iev, dt, "meta", "", "timesig: " + "{0}/{1}, {2} clk/tick, {3} qn/32n".format(nn, 1 << dd, cc, bb))
                            elif metatype == 0x59:
                                sf = reads1(f)
                                mi = readu1(f)
                                settablerow(table, iev, dt, "meta", "", "keysig: " + "{0} ({1} {2})".format(resolvekeyname(sf, mi), sf, "min" if mi else "maj"))
                            else:
                                metadata = f.read(metalen)
                                settablerow(table, iev, dt, "meta", "", "type={0:02x} length={1}: ".format(metatype, metalen) + "".join(["{0:02x} ".format(b) for b in metadata[:128]]) + metaelp)
                        else:
                            settablerow(table, iev, dt, "invalid({0:02x})".format(s), "", "")
                            # f.seek(trkend)
                        iev = iev + 1
                    table.setRowCount(iev)
        except SMFError as e:
            QMessageBox.warning(self, "loadsmf", e.message, QMessageBox.Ok)
        except Exception as e:
            QMessageBox.warning(self, "loadsmf", traceback.format_exc(), QMessageBox.Ok)
    def exportsyx(self):
        if len(self.pathlabel.text()) == 0:
            return
        try:
            midpath = self.pathlabel.text()
            dir = QFileDialog.getExistingDirectory(self, "target folder", os.path.dirname(midpath))
            if 0 < len(dir):
                isyx = -1
                for syx in self.sysexes:
                    if 0 < len(syx):
                        newfile = syx[0] == 0xf0
                        if newfile:
                            isyx = isyx + 1
                        fn = os.path.splitext(os.path.basename(midpath))[0] + "#{0}.syx".format(isyx)
                        syxpath = os.path.join(dir, fn)
                        with open(syxpath, "wb" if newfile else "ab") as f:
                            f.write(syx)
        except Exception as e:
            QMessageBox.warning(self, "exportsyx", traceback.format_exc(), QMessageBox.Ok)

def main():
    app = QApplication(sys.argv)
    mainwnd = MainWindow()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
