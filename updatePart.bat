@echo off

echo Updating paper notes and generating summary file partially...
echo.

::执行Note文件生成操作
python updatePart.py

::切换到writing目录下
cd ..
cd ..
cd F:\zhaoxuhui.github.io
cd writing

::执行expression和word的自动生成
python tool_gen_py3.py

echo.
echo ------------------------------------------
echo All papers are updated and saved in file!
echo All expressions and words are summarized!
echo ------------------------------------------
echo.

pause