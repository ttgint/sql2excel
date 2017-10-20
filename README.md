# sql2excel
Once you have configured your properties files, export your queries into excel sheets.


settings.properties example
```
#Sheet1
SELECT * FROM V$SESSION;

#Sheet2
SELECT * FROM DUAL;
SELECT * FROM V$SESSION;
SELECT * FROM DUAL;

#Sheet3
SELECT * FROM V$SESSION;
```
