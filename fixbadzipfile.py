def fixBadZipfile(zipFile):  
 f = open(zipFile, 'r+b')  
 data = f.read()  
 pos = data.find('\x50\x4b\x05\x06') # End of central directory signature  
 if (pos > 0):  
     self._log("Trancating file at location " + str(pos + 22)+ ".")  
     f.seek(pos + 22)   # size of 'ZIP end of central directory record' 
     f.truncate()  
     f.close()  

fixBadZipfile('request.zip')
