
# import cx_Oracle
# cx_Oracle.init_oracle_client(lib_dir=r"C:\oracle\instantclient_19_11")



import cx_Oracle

cx_Oracle.init_oracle_client(lib_dir=r"C:\oracle\instantclient_19_11")


connection = cx_Oracle.connect(user='ora01', password='oracle_4U2021', dsn='edudb1_high')
cursor = connection.cursor()

# cursor.execute(
# '''
# select empno, ename, job, deptno, sal
# from emp e
# where sal >= (select avg(sal)
#                 from emp
#                 where deptno = e.deptno)''')

# qdata = [empno for empno in cursor]

# qdata = [empno for empno in cursor]

# empno = [x[0] for x in qdata]
# ename = [x[1] for x in qdata]
# job = [x[2] for x in qdata]

# cursor.execute("create table pytab (id number, data varchar2(20))")

# rows = [ (1, "First" ),
#          (2, "Second" ),
#          (3, "Third" ),
#          (4, "Fourth" ),
#          (5, "Fifth" ),
#         #  (6, "Sixth" ),
#          (7, "Seventh" ) ]

# cursor.executemany("insert into pytab (id, data) values (:1, :2)", rows)
# connection.commit()
result = cursor.execute('select * from maillog')
print(result.fetchall())

# for student student_list:
#         cursor.execute(f" insert into student_list (id, name, ..)")values student_id_seq.nextval, 
