Option Explicit

' Checkout Peewee and SQLAlchemy for examples of SQL ORM mappers

	' Table Object (name, Column(...), Column(...), ...)

	' result = conn.execute(ins)

	' s = select([users])
	' result = conn.execute(s)

	' select([users, addresses]).where(users.c.id == addresses.c.user_id)

	

	' insert()
	' insert.values()
	' values
	' from(...)
	' where(...).values(...)
	' and(...)
	' or(...)
	' union(...)
	' union_all(...)
	' .asc() 
	' .desc()
	' count(...)
	' average(...)
	' sum(...)
	' group_by(...)
	' order_by(...)
	' .limit(1).offset(1)
	' having(...)
	' .distinct()
	' .like(...)
	' .join() 
	' .outerjoin() 
	' .IsNull()
	' .IsNotNull()
	' .Top()
	' .Min()
	' .Max()
	' update()
	' update().values(...)
	' delete()
	' create_engine
	' case(...)
	' INSERT INTO SELECT / SELECT INTO
	' ANY / ALL
	' EXISTS
	' BETWEEN

Class base_DB_SQL
	Private Sub Class_Initialize()

	End Sub

	Private Sub Class_Terminate()

	End Sub
End Class

If WScript.ScriptName = "base_DB_SQL.vbs" Then

End If
