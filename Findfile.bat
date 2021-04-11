@echo off
cd /d Â·¾¶
for /r %%i in (Gpoint*.ta) do ( echo %%i > >"GPoint.txt")
for /r  %%i in ( Routing.la*) do ( echo %%i >>"Routing.txt" )
for /r  %%i in ( Boundary.la*) do ( echo %%i >>"Boundary.txt" )
for /r  %%i in ( Sample.ta*) do ( echo %%i >>"Sample.txt" )
for /r  %%i in ( Attitude.ta*) do (echo %%i >>"Attitude.txt" )
for /r  %%i in ( L*.db) do (echo %%i >>RouteDSC.txt )