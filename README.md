# DatabaseMatching
A basic Python script to match a list of numbers to database, both in .xlsx format. This was written freelance for a telecom employee who reached out to me.
Numbers are being checked for correctness and loaded into a Pandas dataframe. The matching databases are also loaded into Pandas dataframes and concatenated.
Using a PhoneNumber dataclass, all relevant attributes of a number are tracked and the numbers are saved in one of three categories:
- found in the database
- missing in the database 
- incorrectly formatted number.

The results are exported as a csv with five columns: 
- number (the number)
- correct (if the number formatted correctly)
- found (was the number found in a database)
- table (in which database was the number found)
- group (is the number part of a larger group of phonenumbers)
- vms (what is the current flag of the number).
