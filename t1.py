city = ['沈阳', '大连', '鞍山', '抚顺', '本溪', '丹东', '锦州', '营口', '阜新', '辽阳', '铁岭', '朝阳', '盘锦', '葫芦岛', '行政审批局']


for i in city:
    print('----------------------------------------------')
    print('''SELECT
            s."车牌号码",
            s."所属行业",
            s."所属企业",
            s."超速报警数",
            ( SELECT sum( s."超速报警数" ) FROM Violations_alarm_top_ten s WHERE s."所属地区" LIKE '%{0}%' and s."所属行业" in ('包车客运','班车客运') )  as 报警总次数
        FROM
            Violations_alarm_top_ten s 
        WHERE
            s."所属地区" LIKE '%{0}%'  and s."所属行业" in ('包车客运','班车客运')
        ORDER BY
            s."超速报警数" + 0 DESC 
            LIMIT 10;
        
        SELECT
            s."车牌号码",
            s."所属行业",
            s."所属企业",
            s."疲劳驾驶报警数",
            ( SELECT sum( s."疲劳驾驶报警数" ) FROM Violations_alarm_top_ten s WHERE s."所属地区" LIKE '%{0}%' and s."所属行业" in ('包车客运','班车客运'))  as 疲劳驾驶报警总数
        FROM
            Violations_alarm_top_ten s 
        WHERE
            s."所属地区" LIKE '%{0}%'  and s."所属行业" in ('包车客运','班车客运')
        ORDER BY
            s."疲劳驾驶报警数" + 0 DESC 
            LIMIT 10;
            
            SELECT
            s."车牌号码",
            s."所属行业",
            s."所属企业",
            s."超速报警数",
            ( SELECT sum( s."超速报警数" ) FROM Violations_alarm_top_ten s WHERE s."所属地区" LIKE '%{0}%' and s."所属行业" in ('危货运输'))  as 报警总次数
        FROM
            Violations_alarm_top_ten s 
        WHERE
            s."所属地区" LIKE '%{0}%'  and s."所属行业" in ('危货运输')
        ORDER BY
            s."超速报警数" + 0 DESC 
            LIMIT 10;
        
        SELECT
            s."车牌号码",
            s."所属行业",
            s."所属企业",
            s."疲劳驾驶报警数",
            ( SELECT sum( s."疲劳驾驶报警数" ) FROM Violations_alarm_top_ten s WHERE s."所属地区" LIKE '%{0}%' and s."所属行业" in ('危货运输'))  as 疲劳驾驶报警总数
        FROM
            Violations_alarm_top_ten s 
        WHERE
            s."所属地区" LIKE '%{0}%'  and s."所属行业" in ('危货运输')
        ORDER BY
            s."疲劳驾驶报警数" + 0 DESC 
            LIMIT 10;'''.format(i))

