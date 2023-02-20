import pysftp

# create connection options object and ignore known_hosts check
cnopts = pysftp.CnOpts()
cnopts.hostkeys = None

# connection parameters
sftpHost = '69.44.220.253'
sftpPort = 22
uname = 'mdo'
pwd = 'kZF0NxLP'

# establish connection
with pysftp.Connection(sftpHost, port=sftpPort, username=uname, password=pwd, cnopts=cnopts) as sftp:
    print("connected to sftp server!")
    sftp.cwd("/Outbound")
    print(sftp.pwd)
    latest = 0
    latestfile = None

    for fileattr in sftp.listdir_attr():
        if fileattr.filename.startswith('Innventures_week_') and fileattr.st_mtime > latest:
            latest = fileattr.st_mtime
            latestfile = fileattr.filename

    if latestfile is not None:
        sftp.get(latestfile, latestfile)