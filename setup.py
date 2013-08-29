__author__ = 'Jason Vanzin'
from cx_Freeze import setup, Executable

base = "Win32GUI"
includefiles = ['icon.png', 'logo.png']
setup(
        name = "Server Disk Space Collector",
        version = "0.1",
        description = "Server Disk Space Collector",

        options = {"build_exe":{"include_msvcr": True, 'include_files':includefiles}},
        executables = [Executable("sdsc.pyw", base=base, icon="icon.ico")]
)




