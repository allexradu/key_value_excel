import platform
import os

table_location = os.path.join('txt', 'sql.txt')

replace_string = '),('
replace_with = '),\n('

with open(table_location, "r+") as f:
    old = f.read()  # read everything in the file
    new = old.replace(replace_string, replace_with)
    print(new)
    f.seek(0)
    f.write(new)
