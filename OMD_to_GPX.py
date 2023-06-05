
import os, sys, datetime, pandas as pd
import xml.etree.ElementTree as ET
from tkinter.filedialog import askdirectory
#Inspired from http://uncledens.chez-alice.fr/python/montregps.htm

CURRENT_DIR = os.path.dirname(sys.argv[0]) #Path of current script
GMT = "+02:00"
DELETE = True

def choose_dir(mydir:str=CURRENT_DIR, mytitle:str='Choose directory containing OMD files')->str:
    return askdirectory(initialdir=mydir, title=mytitle)

class Record(): #Contains a single track point record

    def __init__(self, byte_coord:bytearray, byte_kpi:bytearray):
        self.lat = Record._from_byte(byte_coord[0:4]) / 1000000
        self.lon = Record._from_byte(byte_coord[4:8]) / 1000000
        self.dist = Record._from_byte(byte_coord[8:12])
        self.tps = Record._from_byte(byte_coord[12:14])
        #self.alt = Record._from_byte(byte_coord[15:17])
        self.vit = Record._from_byte(byte_kpi[2:4]) / 100
        self.kCal = Record._from_byte(byte_kpi[4:6])
        self.HR = byte_kpi[6]

    def set_date(self, start_date:datetime.datetime):
        self.date = start_date + datetime.timedelta(seconds=self.tps)

    @classmethod
    def _from_byte(cls, bytearr:bytearray): #takes an array of bytes in reverse ('little') order, and convert it
        return int.from_bytes(bytearr, byteorder='little')

    def __repr__(self):
        return f"Record: {self.__dict__}"

class Records(): #Lits all "Record"s

    def __init__(self, file_full_path:str):
        self.file_full_path = file_full_path
        self.file_path, self.file_name = os.path.split(self.file_full_path)
        self.file_name, self.file_ext = os.path.splitext(self.file_name)
        self.date = datetime.datetime.fromtimestamp(os.path.getmtime(self.file_full_path)) #End date of the recording
        self.file_new_name = os.path.join(self.file_path, "OnMove200_"+self.date.strftime('%Y%m%d-%H%M'))
        self.records_list = self._file_to_records()

    def _file_to_records(self)->list:
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
        self.start_date = self.date - datetime.timedelta(seconds=records_list[-1].tps)
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
                if getattr(rec, 'alt', None):
                    elem = ET.SubElement(trkpt, 'ele')
                    elem.text = str(rec.alt)
                elem = ET.SubElement(trkpt, 'time')
                elem.text = rec.date.strftime('%Y-%m-%dT%H:%M:%S')+GMT
                ext = ET.SubElement(trkpt, 'extensions')
                trackpoint = ET.SubElement(ext, f"{{{ns_gpxtpx}}}TrackPointExtension")
                elem = ET.SubElement(trackpoint, f"{{{ns_gpxtpx}}}hr")
                elem.text = str(rec.HR)
            tree.write(xml_file, encoding='UTF-8', xml_declaration=True, method='xml')

    def save_to_csv(self): #Export in GPX (XML) format
        df = pd.DataFrame([rec.__dict__ for rec in self.records_list])
        df.to_csv(self.file_new_name+'.csv', sep='\t', index=False)

    def __repr__(self):
        return f"Records: {len(self.records_list)} measures recorded on {self.date:%Y-%m-%d %H:%M}"

def main():
    dir_path = choose_dir()
    if dir_path:
        _, _, filenames = next(os.walk(dir_path))
        for f in filenames:
            if os.path.splitext(f)[1].upper() == '.OMD':
                print(f"Exporting {f}")
                records = Records(os.path.join(dir_path, f))
                records.save_to_gpx()
                #records.save_to_csv()
                if DELETE:
                    os.remove(os.path.join(dir_path, f))
                    OMH = os.path.join(dir_path, f[:-4]+".OMH")
                    if os.path.exists(OMH):
                        os.remove(OMH)

if __name__ == '__main__':
    main()

#Possible enhancements: Get altitude, Get OpenStreetMap,...
#See:
    #https://github.com/ColinPitrat/kalenji-gps-watch-reader
    #https://github.com/tumic0/GPXSee
    #https://www.outofpluto.com/blog/get-out-of-the-decathlon-geonaute-jail/
