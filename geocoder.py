##pyodbc and omgeo are required.
## >pip install pyodbc
## >pip install omgeo


import pyodbc,csv
from omgeo import Geocoder

#build Geocoder instance
g = Geocoder()

#Connect to Falcon using SQL Server Native Client Connection String
sqlcon =   pyodbc.connect("Driver={SQL Server Native Client 10.0};Server=FALCON;Database=student_view;Trusted_Connection=yes;")
sqlcur = sqlcon.cursor()

#Query that will be passed to the cursor
query = """SELECT cec_key,addr_1,city,state,zip5
        FROM client_entering_class
        WHERE
        school_key = ?
        AND entry_year = ?
        AND app_stat_ind = 1
        """

#Containers for data
students = []
errors = []

#Define fixNull function for data cleaning
def fixNull(value):
    if value is None:
        return " "
    else:
        return value


#Fetch all results from SQL cursor
results = sqlcur.execute(query,1067,2015).fetchall()
print 'student records returned by SQL cursor'


#Loop through results and build dictionary object
for s in results:
    paddress = str(s.addr_1)+', '+str(s.city)+' '+str(fixNull(s.state))+', '+str(fixNull(s.zip5))
    student = {
        "cec_key": s.cec_key,
        "addr_1" : s.addr_1,
        "city" : s.city,
        "state" : s.state,
        "zip" : str(s.zip5),
        "Proper Address" : paddress,
        "Latitude" : '',
        "Longitude" : ''}
    students.append(student)

print 'records now exist in "students" list'


#Looping through list of dictionaries and adding geocoded results
## Default geocoding service is EsriWGS
i = len(students)
print '%s records remaining' % str(i)
for s in students:
    i -= 1
    gcode = g.geocode(s["Proper Address"])
    try:
        s["geocode response"] = gcode["candidates"][0]
        s["Latitude"] = gcode["candidates"][0].y
        s["Longitude"] = gcode["candidates"][0].x
        print 'Good: %s records remaining' % str(i)
    except KeyboardInterrupt:
        break
    except:
        errors.append(s)
        print 'Bad: %s records remaining' % str(i)
        continue

print 'Geocoding Complete'



#Write output to .csv file
with open('c:\\users\\kvanderbush\\desktop\\quinLatLong.csv','wb') as file:
    writer = csv.writer(file)
    writer.writerow(('cec_key','Proper Address','Latitude','Longitude'))
    for s in students:
        try:
            row = (s['cec_key'],s['Proper Address'],fixNull(s['Latitude']),fixNull(s['Longitude']))
            writer.writerow(row)
        except:
            pass
print '%i records processed with %i errors' % (len(students),len(errors))
