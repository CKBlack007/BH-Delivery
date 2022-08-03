###
This code was written to drive Delivery functions for a large catering organization. They work in coordination with Google sheets. Here is the prototype sheet
https://docs.google.com/spreadsheets/d/1hzq3orJPM74VIYZSaDd0FHDjX9tRjlefDscEvEwh1Ic/edit?usp=sharing

CTE	- wrapper for functions should the Compare Travel Times function was completed. orphaned
AuditTimes- copies current record of travel times to rTimeLog. Designed to run periodically. orphaned
CompareTimes-start of function to compare  travel times between hours. a small function that was to be part of a larger program that would alert me to travel time estimate changes which was never completed
TravelTime-Gets travel time estimates between destinations and applies them to their respective entries in the Engine worksheet
BookName-makes the workbook filename variable accessible within the workbook
cleanCalendarCreator