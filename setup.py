from cx_Freeze import setup, Executable 
import sys 

buildOptions = dict(packages = ["openpyxl"],  

excludes = []) 
exe = [Executable("main.py")] 

setup(  
    name= 'frontier_spliter',  
    version = '0.1',  
    author = "Kimshin",  
    description = "frontier people spliter ",  
    options = dict(build_exe = buildOptions),  
    executables = exe 
    )
