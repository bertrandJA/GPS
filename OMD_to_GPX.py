
import os, sys, datetime, pandas as pd
import xml.etree.ElementTree as ET
from tkinter.filedialog import askdirectory
#Inspired from http://uncledens.chez-alice.fr/python/montregps.htm
#Format of OMD and OMH from https://github.com/ylecuyer/OnMove200/blob/master/main.rb

CURRENT_DIR = os.path.dirname(sys.argv[0]) #Path of current script
GMT = "+02:00"
DELETE = False
BACKUP = True

def choose_dir(mydir:str, mytitle:str='Choose directory containing OMD files')->str:
    return askdirectory(initialdir=mydir, title=mytitle)

def split_path_ext(full_path:str):
    path, filename = os.path.split(full_path)
    filename, ext = os.path.splitext(filename)
    return path, filename, ext

class OMH(): #Contains a summary of the records

    def __init__(self, file_full_path:str): #Decode OMH file
        self.file_full_path = file_full_path
        self.path, self.filename, self.ext = split_path_ext(file_full_path)
        with open(self.file_full_path, 'rb') as omd:
            bytearr = omd.read() #60 bytes
            self.distance = Record._bytes_to_int(bytearr[0:4]) / 1000 #in km
            self.duration = Record._bytes_to_int(bytearr[4:6]) #sec
            self.avgSpeed = Record._bytes_to_int(bytearr[6:8]) / 100 #km/s
            self.maxSpeed = Record._bytes_to_int(bytearr[8:10]) / 100 #km/s
            self.totalKcal = Record._bytes_to_int(bytearr[10:12])
            self.avgHR = bytearr[12]  #Heart rate in bpm
            self.maxHR = bytearr[13]
            self.year = bytearr[14] + 2000 #bytes store last 2 digits. Ex: 23 for 2023
            self.month = bytearr[15]
            self.day = bytearr[16]
            self.hour = bytearr[17]  #arrival time
            self.mn = bytearr[18]
##            self.actID1 = bytearr[19]
##            self.pointsCount = Record._bytes_to_int(bytearr[20:22])
##            self.indoor = bytearr[22]
##            self.buffer1 = Record._bytes_to_int(bytearr[23:38])
##            self.actID2 = bytearr[38]
##            self.dataID2 = bytearr[39]
##            self.timeBelow = Record._bytes_to_int(bytearr[40:42])
##            self.timeIn = Record._bytes_to_int(bytearr[42:44])
##            self.timeAbove = Record._bytes_to_int(bytearr[44:46])
##            self.speedLimitLow = Record._bytes_to_int(bytearr[46:48])
##            self.speedLimitHigh = Record._bytes_to_int(bytearr[48:50])
            self.hrLimitLow = bytearr[50] #70 instead of 60??
            self.hrLimitHigh = bytearr[51]  #173 instead of 176??
##            self.target = bytearr[52]
##            self.buffer1 = Record._bytes_to_int(bytearr[53:58])
##            self.actID3 = bytearr[58]
##            self.dataID3 = bytearr[59]

    def __repr__(self):
        return f"OMH: {self.__dict__}"

class Record(): #Contains a single track point record

    def __init__(self, byte_coord:bytearray, byte_kpi:bytearray):   #Decode record
        self.lat = Record._bytes_to_int(byte_coord[0:4]) / 1000000
        self.lon = Record._bytes_to_int(byte_coord[4:8]) / 1000000
        self._dist = Record._bytes_to_int(byte_coord[8:12])
        self._sec = Record._bytes_to_int(byte_coord[12:14]) #seconds elapsed since the start
        #self.ele = Record._bytes_to_int(byte_coord[15:17])    #elevation
        self._speed = Record._bytes_to_int(byte_kpi[2:4]) / 100
        self.kCal = Record._bytes_to_int(byte_kpi[4:6])
        self.HR = byte_kpi[6]   #Heart Rate

    def set_date(self, start_date:datetime.datetime):
        self.date = start_date + datetime.timedelta(seconds=self._sec)

    @classmethod
    def _bytes_to_int(cls, bytearr:bytearray)->int: #takes an array of bytes in reverse ('little') order, and convert it
        return int.from_bytes(bytearr, byteorder='little')

    def __repr__(self):
        return f"Record: {self.__dict__}"

class Records(): #Lits of all "Record"s

    def __init__(self, file_full_path:str):
        self.file_full_path = file_full_path
        self.path, self.filename, self.ext = split_path_ext(self.file_full_path)
        self.omh_path = os.path.join(self.path, self.filename) + '.OMH'
        if os.path.exists(self.omh_path):   #Read OMH file to extract date
            self.omh = OMH(self.omh_path)
            self.date = datetime.datetime(self.omh.year, self.omh.month, self.omh.day, self.omh.hour, self.omh.mn, 00)
            print(self.omh)
        else:   #No OMH file available. We will consider that date is the modif time of OMD file
            self.omh_path = None
            self.date = datetime.datetime.fromtimestamp(os.path.getmtime(self.file_full_path)) #End date of the recording
            #self.date = datetime.datetime(2023, 6, 12, 20, 25, 00) #Uncomment if you want to force date
        self.file_new_name = os.path.join(self.path, "OnMove200_"+self.date.strftime('%Y%m%d-%H%M'))
        self.records_list = self._file_to_records() #Collect each individal Record

    def _file_to_records(self)->list:   #Collect each individal Record
        records_list = []
        with open(self.file_full_path, 'rb') as omd:
            bytearr = omd.read(60)
            while bytearr: #loop on chunks of 60 bytes, corresponding to 2 records
                if len(bytearr)==60:
                    records_list += [ Record(bytearr[0:20], bytearr[40:50]) ]
                    records_list += [ Record(bytearr[20:40], bytearr[50:60]) ]
                else: #last chunk may contain only 40 bytes, corresponding to 1 record
                    records_list += [ Record(bytearr[0:20], bytearr[20:30]) ]
                bytearr = omd.read(60)
        #Determine start date of the recording
        self.start_date = self.date - datetime.timedelta(seconds=records_list[-1]._sec)
        for rec in records_list: #Update all records, based on start date
            rec.set_date(self.start_date)
        return records_list

    def save_to_gpx(self): #Export in GPX (XML) format
        ns_default = "http://www.topografix.com/GPX/1/1"
        #ns_xsi = "http://www.w3.org/2001/XMLSchema-instance"
        ns_gpxtpx = "http://www.garmin.com/xmlschemas/TrackPointExtension/v1"
        with open(self.file_new_name+'.gpx', mode='wb') as xml_file: #encoding='utf-8'
            ET.register_namespace("", ns_default)
            #ET.register_namespace("schemaLocation", f"{ns_default} {ns_default}/gpx.xsd {ns_gpxtpx} {ns_gpxtpx}.xsd")
            #ET.register_namespace("xsi", ns_xsi)
            ET.register_namespace("gpxtpx", ns_gpxtpx)
            root = ET.Element(f"{{{ns_default}}}gpx")
            root.set('version', '1.1')
            root.set('creator', "ONConnect")
            tree = ET.ElementTree(root)
            meta = ET.SubElement(root, 'metadata')
            elem = ET.SubElement(meta, 'name')
            elem.text = "ONmove 200 " + self.date.strftime('%Y-%m-%d %H:%M')
            #elem = ET.SubElement(meta, 'time')
            #elem.text = self.date.strftime('%Y-%m-%dT%H:%M:%S')
            trk = ET.SubElement(root, 'trk')
            #elem = ET.SubElement(trk, 'name')
            #elem.text = "ONmove 200 "  + self.date.strftime('%Y-%m-%d %H:%M')
            trkseg = ET.SubElement(trk, 'trkseg')
            for rec in self.records_list:
                trkpt = ET.SubElement(trkseg, 'trkpt', attrib={'lat': str(rec.lat), 'lon': str(rec.lon)})
                if getattr(rec, 'ele', None):
                    elem = ET.SubElement(trkpt, 'ele')
                    elem.text = str(rec.alt)
                elem = ET.SubElement(trkpt, 'time')
                elem.text = rec.date.strftime('%Y-%m-%dT%H:%M:%S')+GMT
                ext = ET.SubElement(trkpt, 'extensions')
                trackpoint = ET.SubElement(ext, f"{{{ns_gpxtpx}}}TrackPointExtension")
                elem = ET.SubElement(trackpoint, f"{{{ns_gpxtpx}}}hr")
                elem.text = str(rec.HR)
            tree.write(xml_file, encoding='UTF-8', xml_declaration=True, method='xml')

    def save_to_csv(self): #Export to csv
        df = pd.DataFrame([rec.__dict__ for rec in self.records_list])
        df.to_csv(self.file_new_name+'.csv', sep='\t', index=False)

    def __repr__(self):
        return f"Records: {len(self.records_list)} measures recorded on {self.date:%Y-%m-%d %H:%M}"

def main():
    dir_path = choose_dir(CURRENT_DIR)
    if dir_path:
        _, _, filenames = next(os.walk(dir_path)) #List all files in dir_path
        for f in filenames: #loop on each file
            if os.path.splitext(f)[1].upper() == '.OMD':
                print(f"Exporting {f}")
                records = Records(os.path.join(dir_path, f))    #Decode records in file
                records.save_to_gpx()   #Create a .GPX file
                #records.save_to_csv()
                if BACKUP:  #Move OMH and OMD files to a subdirectory 'Save'
                    os.rename(os.path.join(dir_path, f), os.path.join(dir_path, 'Save', f))
                    if records.omh_path is not None:
                        os.rename(records.omh_path, os.path.join(dir_path, 'Save', records.omh.filename)+records.omh.ext)
                elif DELETE:    #Delete OMD and OMH files
                    os.remove(os.path.join(dir_path, f))
                    if records.omh_path is not None:
                        os.remove(records.omh_path)

if __name__ == '__main__':
    main()

#Possible enhancements: Get altitude, Get OpenStreetMap,...
#See:
    #https://github.com/ColinPitrat/kalenji-gps-watch-reader
    #https://github.com/tumic0/GPXSee
    #https://www.outofpluto.com/blog/get-out-of-the-decathlon-geonaute-jail/
