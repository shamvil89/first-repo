declare @reginfo table (value varchar(50), Data nvarchar(50))
insert into @reginfo  select 'servername', @@SERVERNAME

insert into @reginfo  exec xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\SuperSocketNetLib\Tcp',
@value_name = 'DisplayName' 

insert into @reginfo exec xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\SuperSocketNetLib\Tcp',
@value_name = 'Enabled'

insert into @reginfo  exec xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\SuperSocketNetLib\Tcp\IpAll',
@value_name = 'TcpPort' 

insert into @reginfo EXEC xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\SuperSocketNetLib\Tcp\IpAll',
@value_name = 'TcpDynamicPorts'

insert into @reginfo EXEC xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\SuperSocketNetLib\Np',
@value_name = 'DisplayName'

insert into @reginfo EXEC xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\SuperSocketNetLib\Np',
@value_name = 'Enabled'

insert into @reginfo EXEC xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\SuperSocketNetLib\Np',
@value_name = 'PipeName'

insert into @reginfo EXEC xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\SuperSocketNetLib\Sm',
@value_name = 'DisplayName'
insert into @reginfo EXEC xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\SuperSocketNetLib\Sm',
@value_name = 'Enabled'

insert into @reginfo select 'HADR related info', '------'

insert into @reginfo EXEC xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\HADR',
@value_name = 'HADR_Enabled'
insert into @reginfo select 'ErrorLog file location', '------'

insert into @reginfo EXEC xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\Parameters',
@value_name = 'SQLArg1'
insert into @reginfo select 'Filestream', '------'

insert into @reginfo EXEC xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\Filestream',
@value_name = 'Enablelevel'

insert into @reginfo EXEC xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\SuperSocketNetLib',
@value_name = 'DisplayName'
insert into @reginfo EXEC xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\SuperSocketNetLib',
@value_name = 'ForceEncryption'
insert into @reginfo EXEC xp_instance_regread
@rootkey = 'HKEY_LOCAL_MACHINE',
@key =
'Software\Microsoft\Microsoft SQL Server\MSSQLServer\SuperSocketNetLib',
@value_name = 'HideInstance'


select * from @reginfo where data is not null

