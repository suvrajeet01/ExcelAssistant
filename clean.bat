@echo off

echo "This will remove all directories and files created by building/running the program"
pause

rmdir __pycache__ /s /q
rmdir progress\__pycache__ /s /q
rmdir dist /s /q
rmdir build /s /q
del /q main.spec

pause