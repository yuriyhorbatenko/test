export const example = `#group,false,false,true,true,false,false,true,true,true,true
#datatype,string,long,dateTime:RFC3339,dateTime:RFC3339,dateTime:RFC3339,double,string,string,string,string
#default,_result,,,,,,,,,
,result,table,_start,_stop,_time,_value,_field,_measurement,host,region
,,0,2022-12-31T05:41:24Z,2023-01-31T05:41:24.001Z,2023-01-01T00:00:00Z,15.43,mem,m,A,east
,,1,2022-12-31T05:41:24Z,2023-01-31T05:41:24.001Z,2023-01-01T00:00:00Z,59.25,mem,m,B,east
,,2,2022-12-31T05:41:24Z,2023-01-31T05:41:24.001Z,2023-01-01T00:00:00Z,52.62,mem,m,C,east

#group,false,false,true,true,true,true,false,false,true,true
#datatype,string,long,string,string,dateTime:RFC3339,dateTime:RFC3339,dateTime:RFC3339,string,string,string
#default,_result,,,,,,,,,
,result,table,_field,_measurement,_start,_stop,_time,_value,host,region
,,3,mem_level,m,2022-12-31T05:41:24Z,2023-01-31T05:41:24.001Z,2023-01-01T00:00:00Z,ok,A,east
,,4,mem_level,m,2022-12-31T05:41:24Z,2023-01-31T05:41:24.001Z,2023-01-01T00:00:00Z,info,B,east
,,5,mem_level,m,2022-12-31T05:41:24Z,2023-01-31T05:41:24.001Z,2023-01-01T00:00:00Z,info,C,east`;
