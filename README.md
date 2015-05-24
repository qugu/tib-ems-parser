# tib-ems-parser
A Perl config file parser for TIBCO EMS's tibemsd.json to ease user and topic revalidation.
At this point, the script exports data from a JSON file to XML for general usage.

Plans:
* Use database (MongoDB?) to store and read old previous revalidation data;
* Make user CLI to find JSON and ask user what he wants to do;
* Comparison of old and new user/destination lists and export result for reconcillation;

