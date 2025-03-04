from struct import unpack
from io import BytesIO, StringIO
import math
import binascii
import json
import csv
from schema import PidTagSchema

def hexify(PropID):
    return "{0:#0{1}x}".format(PropID, 10).upper()[2:]

class OABParser:
    def __init__(self, data=None):
        self.data = data
        self.ulVersion = None
        self.ulSerial = None
        self.ulTotRecs = None
        self.HDR_cAtts = None
        self.OAB_cAtts = None
        self.OAB_Atts = []
        self.all_records = []
        self.parsed = False

    def parse(self, data=None):
        """Parses the OAB data.  If data is not provided in the constructor, it must be provided here."""

        if data is None and self.data is None:
            raise ValueError("No OAB data provided.")
        
        if data:
            self.data = data

        f = BytesIO(self.data)
        self._parse_header(f)  # Parse the OAB header
        self._parse_metadata(f)  # Parse metadata sections

        counter = 0
        max_records = 200000

        while counter < max_records:
            counter += 1
            try:
                record = self._parse_record(f)  # Parse a single record
                if record:
                    self.all_records.append(record)
                else:
                    break  # No more records
            except Exception as e:
                break

        self.parsed = True
        return self

    def _parse_header(self, f):
        """Parses the OAB header."""
        (self.ulVersion, self.ulSerial, self.ulTotRecs) = unpack('<III', f.read(4 * 3))
        assert self.ulVersion == 32, 'This only supports OAB Version 4 Details File'

    def _parse_metadata(self, f):
        """Parses metadata sections of the OAB file."""
        cbSize = unpack('<I', f.read(4))[0]
        meta = BytesIO(f.read(cbSize - 4))

        self.HDR_cAtts = unpack('<I', meta.read(4))[0]
        for _ in range(self.HDR_cAtts):
            meta.read(8)  # Skip ulPropID and ulFlags

        self.OAB_cAtts = unpack('<I', meta.read(4))[0]
        self.OAB_Atts = []
        for _ in range(self.OAB_cAtts):
            ulPropID = unpack('<I', meta.read(4))[0]
            meta.read(4)  # Skip ulFlags
            self.OAB_Atts.append(ulPropID)

        cbSize = unpack('<I', f.read(4))[0]
        f.read(cbSize - 4)

    def _parse_record(self, f):
        """Parses a single OAB record."""
        read = f.read(4)
        if len(read) == 0 or len(read) < 4:
            return None  # Indicate end of records
        
        cbSize = unpack('<I', read)[0]
        chunk = BytesIO(f.read(cbSize - 4))
        presenceBitArray = bytearray(chunk.read(int(math.ceil(self.OAB_cAtts / 8.0))))
        indices = [i for i in range(self.OAB_cAtts) if (presenceBitArray[i // 8] >> (7 - (i % 8))) & 1 == 1]

        rec = {}
        for i in indices:
            PropID = hexify(self.OAB_Atts[i])
            if PropID not in PidTagSchema:
                continue #Skip if PropID not in schema

            try:
                (Name, Type) = PidTagSchema[PropID]
                val = self._read_property(chunk, Name, Type)
                rec[Name] = val
            except Exception as e:
                print(f"Error reading property {PropID}: {e}")
                # Optionally, re-raise the exception or return None, depending on desired behavior
                continue

        chunk.read()  # Ensure chunk is fully read to prevent issues.
        return rec

    def _read_property(self, chunk, Name, Type):
        """Reads a property value from the chunk based on its type."""

        def read_str():
            buf = b""
            while True:
                n = chunk.read(1)
                if n == b"\0" or n == b"":
                    break
                buf += n
            try:
                val = buf.decode('utf-8')
            except UnicodeDecodeError:
                val = buf.decode('latin-1')
            return val

        def read_int():
            read = chunk.read(1)
            if len(read) < 1:
                return -1
            byte_count = unpack('<B', read)[0]
            if 0x81 <= byte_count <= 0x84:
                bytes_to_read = byte_count - 0x80
                read_data = chunk.read(bytes_to_read)
                if len(read_data) < bytes_to_read:
                    return -1
                byte_count = unpack('<I', (read_data + b"\0\0\0")[0:4])[0]
            else:
                if byte_count > 127:
                    return -1
            return byte_count

        if Type == "PtypString8" or Type == "PtypString":
            return read_str()
        elif Type == "PtypBoolean":
            return unpack('<?', chunk.read(1))[0]
        elif Type == "PtypInteger32":
            return read_int()
        elif Type == "PtypBinary":
            bin_val = chunk.read(read_int())
            return binascii.b2a_hex(bin_val).decode('latin-1')
        elif Type == "PtypMultipleString" or Type == "PtypMultipleString8":
            byte_count = read_int()
            arr = [read_str() for _ in range(byte_count)]
            return arr
        elif Type == "PtypMultipleInteger32":
            byte_count = read_int()
            arr = []
            for _ in range(byte_count):
                val = read_int()
                if Name == "OfflineAddressBookTruncatedProperties":
                    val = hexify(val)
                    if val in PidTagSchema:
                        val = PidTagSchema[val][0]
                arr.append(val)
            return arr
        elif Type == "PtypMultipleBinary":
            byte_count = read_int()
            arr = []
            for _ in range(byte_count):
                bin_len = read_int()
                bin_val = chunk.read(bin_len)
                arr.append(binascii.b2a_hex(bin_val).decode('latin-1'))
            return arr
        else:
            raise ValueError("Unknown property type (" + Type + ")")


    def get_records(self):
        """Returns the parsed records.  Must call parse() first."""
        if not self.parsed:
            raise ValueError("OAB data must be parsed first. Call parse().")
        return self.all_records

    def to_json(self, indent=4):
        """Returns the parsed records as a JSON string.  Must call parse() first."""
        if not self.parsed:
            raise ValueError("OAB data must be parsed first. Call parse().")
        return json.dumps(self.all_records, indent=indent)

    def to_csv(self, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL):
        """Returns the parsed records as a CSV string.  Must call parse() first."""
        if not self.parsed:
            raise ValueError("OAB data must be parsed first. Call parse().")

        if not self.all_records:
            return ""  # Return empty string if no records are found

        # Determine all possible headers from the records.
        headers = set()
        for record in self.all_records:
            headers.update(record.keys())
        headers = sorted(list(headers))  # Sort headers for consistent output

        output = StringIO()  # Use BytesIO to construct CSV string in memory
        writer = csv.writer(output, delimiter=delimiter, quotechar=quotechar, quoting=quoting, lineterminator='\n')

        writer.writerow(headers)  # Write the header row

        for record in self.all_records:
            row = [record.get(header, '') for header in headers]  # Get value or empty string if header doesn't exist
            writer.writerow(row)

        return output.getvalue()


    def save_json(self, filename, indent=4):
         """Saves the parsed records to a JSON file. Must call parse() first."""
         if not self.parsed:
            raise ValueError("OAB data must be parsed first. Call parse().")
         with open(filename, 'w') as json_out:
            json_out.write(json.dumps(self.all_records, indent=indent))

    def save_csv(self, filename, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL):
        """Saves the parsed records to a CSV file.  Must call parse() first."""
        if not self.parsed:
            raise ValueError("OAB data must be parsed first. Call parse().")

        csv_string = self.to_csv(delimiter, quotechar, quoting)
        with open(filename, 'w', newline='', encoding='utf-8') as csv_file:
            csv_file.write(csv_string)

