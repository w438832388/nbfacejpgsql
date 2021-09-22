#encoding=-utf-8
# 蓝本5.3人脸批量导入V3.2（数据库版）（测试时适配蓝本版本20200330）
#暂时只支持jpg格式图片
#需要图片按姓名命名和本程序放入同一文件夹,然后双击执行
#本程序会自动根据图片名称构建人事资料并把人事信息和人脸写入蓝本数据库

#代码如下：
import pyodbc
import os
print('===========================================')
print('======蓝本人脸导入程序 V3.2 by--Wang=======')
print('===========================================')
print('注意事项：')
print('一.支持jpg格式图片导入')
print('二.图片按姓名命名和本程序放入同一文件夹内')
print('三.数据库SA账户密码默认为000000')
print('四.运行时关闭蓝本等其他软件,避免数据冲突')

print('\r\n')

while 1:
        dept='公司'
        ipt=input('请输入部门名称:')
        if ipt=='':
                break
        else:
            dept=ipt
            break

with pyodbc.connect(DRIVER='{SQL Server}',SERVER='(local)',DATABASE='SystemData_MJ',UID='sa',PWD='000000') as db:
    cursor=db.cursor()

    #获取部门信息，有此部门则直接映射部门号，没有则开始执行部门创建
    sql_dept='''SELECT vDepartmentID,vDepartmentName FROM SystemData_MJ.dbo.Department WHERE vDepartmentName=?'''
    cursor.execute(sql_dept,(dept,))#注意这里是元组形式映射赋值，所以只有单个元素时不能忘记逗号
    deptdata=cursor.fetchall()
    if deptdata :
        deptnum=deptdata[0][0]
    else :
        sql_makedept='''INSERT INTO SystemData_MJ.dbo.Department (vDepartmentID,
        vDepartmentName,vDepartmentKey,vDepartmentParent,iLevel) VALUES (?,?,?,?,?)'''

        sql_select_dept='''SELECT iID FROM SystemData_MJ.dbo.Department'''
        cursor.execute(sql_select_dept)
        deptiid=cursor.fetchall()[-1][0]
        deptnum = str(deptiid + 1).zfill(18)
        cursor.execute(sql_makedept,(deptnum,dept,str(deptiid+1),'00140227151649',1))

    #获取人员表的自增长ID以便接下来续写表
    sql_eid='''SELECT EId FROM SystemData_MJ.dbo.Employee'''
    cursor.execute(sql_eid)
    eid=cursor.fetchall()
    if eid:
        num = eid[-1][0]+1
    else:
        num = 1

    #获取目录下后缀为jpg的图片列表，这里的os.listdir()是随机操作，一定要使用list.sort()排序使其结果输出绝对
    jpgname = [item[:-4] for item in os.listdir('./') if item[-4:] == '.jpg']
    jpgname.sort()#不要忘记排序，以使结果绝对化

    #以图片名称来循环构建人员信息表
    for name in jpgname:
        with open(f'{name}.jpg','rb') as jpgfile:
            photo = jpgfile.read()

            uid = f'EMP{num}'
            empid = str(num)
            empname = name
            cardno = f'{num}'.zfill(10)

            sql='''INSERT INTO SystemData_MJ.dbo.Employee(vUID,vEmp_id,vEmp_name,
            vCardNo,vDepart,vDoorPassword,dBeginDate,dEndDate,photo,vEmpSex,dEmpBirthday,
            vEmpEduBackground,vEmpPoliticalBackground,vEmpNation,vEmpProvinceName,
            dEmpWorkIn,dEmpWorkOut,vAttendance) VALUES(?,?,?,?,?,'000000',
            '2000-01-01 00:00:00.000','2035-04-03 00:00:00.000',?,'男','2000-01-01 00:00:00.000',
            '请选择','请选择','汉族','江西省','2017-10-01 00:00:00.000','2035-04-03 00:00:00.000','1')'''

            cursor.execute(sql,(uid,empid,empname,cardno,deptnum,pyodbc.Binary(photo)))
            num+=1

            print(name.ljust(8-len(name))+'导入成功！')

print('\r\n')
print('人脸全部导入成功！')
print('\r\n')

while 1:
        tc=input('按【回车键】退出本程序...')
        if tc=='':
                break
