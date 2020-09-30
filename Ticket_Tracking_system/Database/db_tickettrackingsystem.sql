create database db_tickettracking
use db_tickettracking

create table Employee(Eid varchar(7) primary key,EmployeeName varchar(25),Hire_Date Date,Dept varchar(7))
insert into Employee Values('E100100','venkat',Cast(N'2004-1-10' As Date),'MGM')
insert into Employee Values('E100101','krishna',Cast(N'2004-1-10' As Date),'MGM')
insert into Employee Values('E100102','chandrashekhar',Cast(N'2005-3-11' As Date),'DEV')
insert into Employee Values('E100103','saheer Ali Khan',Cast(N'2008-10-13' As Date),'DEV')
insert into Employee Values('E100104','Shashikanth',Cast(N'2007-2-17' As Date),'DEV')
insert into Employee Values('M100103','Avinash',Cast(N'2007-3-10' As Date),'DEVOPS')
insert into Employee Values('M100105','Ashok',Cast(N'2008-6-18' As Date),'DEVOPS')

create table EmployeeAuthentication(Eid varchar(7) Foreign key references Employee(Eid),UserId varchar(25),userpassword varchar(25))
insert into EmployeeAuthentication Values('E100100','Venkat','Venkat@123')
insert into EmployeeAuthentication Values('E100101','Krishna','Krishna@123')
insert into EmployeeAuthentication Values('E100102','Chandrashekhar','ChandraShekhar@123')
insert into EmployeeAuthentication Values('E100103','Saheer Ali Khan','Saheer@123')
insert into EmployeeAuthentication Values('E100104','Shashikanth','Shashi@123')
insert into EmployeeAuthentication Values('M100103','Avinash','Avinash@123')
insert into EmployeeAuthentication Values('M100105','Ashok','Ashok@123')


Create Table Ticket(Ticket_Id int primary key identity (1,1),Logged_By varchar(7),Raised_Date datetime,Severity varchar(15),Ticket_Desc varchar(30),Resolved_By varchar(7),Resolution varchar(30),Resolved_date DateTime,status varchar(10))
insert into Ticket values('E100101',Cast(N'2012-10-3' As Datetime),'major','App not working','M100103','Need to restart with Lan Cable',Cast(N'2012-10-4' As Datetime),'closed')
insert into Ticket values('E100104',Cast(N'2013-7-10' As Datetime),'critical','Laptop Restart Problem',Null,Null,Null,'Open')

update Ticket set Resolved_By = n
select * from ticket

select * from Employee

select EmployeeName from Employee where Dept <> 'Devops'

Alter procedure sp_CreateTicket(
@Ename varchar(7),
@date varchar(30),
@severity varchar(10),
@Desc varchar(30)
)
As
Begin
	Begin try
		Declare @Eid varchar(7)
		Select @Eid = Eid from Employee where EmployeeName = @Ename
		begin tran
		Insert into Ticket values(@Eid,cast(@date As Datetime),@severity,@Desc,null,null,null,'open')
		commit
	End try
	Begin Catch
		Rollback tran
		raiserror ('Error while Creating Ticket',16,1)
	End Catch
end

Begin
declare @message varchar(30)
exec  sp_CreateTicket 'E10010','2004-01-10','major','jhsvjcjjhc',@message out
select @message
Rollback

delete 
select * from ticket

Alter procedure sp_CloseTicket(
@TicketId int,
@Employee varchar(30),
@Resolution varchar(10),
@status int out
)
As
Begin
	Declare @Eid varchar(7)
	Begin try
		begin tran
		select @Eid = Eid from Employee where EmployeeName = @Employee
		update Ticket set Resolved_By = @Eid ,Resolution = @Resolution ,Resolved_date = GetDate(),[status] = 'close' where Ticket_Id = @TicketId
		Set @status =  @TicketId 
		commit
	End try
	Begin Catch
		Rollback tran
		raiserror ('Error while Closing Ticket',16,1)
	End Catch
end


Begin Tran
declare @Status int
exec  sp_CloseTicket 3,'ashok','jhsvjBcjjhfzdbbsc',@Status out
select @Status
Rollback

Create procedure sp_Report
As
Begin
	Begin try
		Select Logged_By,Ticket_Id,SeverityGETDATE()
		commit
	End try
	Begin Catch
		Rollback tran
		raiserror ('Error while Closing Ticket',16,1)
	End Catch
end


Begin Tran
declare @Status int
exec  sp_CloseTicket 3,'ashok','jhsvjBcjjhfzdbbsc',@Status out
select @Status
Rollback

