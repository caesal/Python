# Imagine you're reading lines from a log file. Each line is a string like this:
 #  	TIMESTAMP=1234567890, EVENT=SOME_EVENT, VALUE=12.5ps
 #  Write a python function that takes one of these strings as input and returns a dictionary like this:
 #  	{'TIMESTAMP': 1234567890, 'EVENT': 'SOME_EVENT', 'VALUE': '12.5ps'}


def get_log(line:str) -> dict:
    parts = [p.strip() for p in line.split(",")]
    output = {}

    for part in parts:
        a, b = part.split("=")
        a = a.strip()
        b = b.strip()

        if a == "TIMESTAMP" :
            try:
                output[a] = int(b)
            except ValueError:
                output = {}
                print(output)
                return output
        else:
            output[a] = b
    print(output)
    return output

line = 'TIMESTAMP=123asv, EVENT=SOME_EVENT, VALUE=12.5ps'
get_log(line)

Imagine you have two SQL tables:
   	unit_info
   	- serial_number (VARCHAR, primary key)
   	- product_sku (VARCHAR)
   	- firmware_version (VARCHAR)
   	test_runs
   	- serial_number (VARCHAR)
   	- test_name (VARCHAR)
   	- timestamp (VARCHAR)
   	- status (VARCHAR)
   Write a SQL query to find the serial_number and the timestamp for all units with the SKU 'SOME-SKU-99' that failed the 'SOME_TEST' test