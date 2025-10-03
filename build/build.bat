CD /D %~dp0
call npm install -g grunt-cli
call npm ci
rem call grunt --level=WHITESPACE_ONLY --desktop=false --formatting=PRETTY_PRINT
rem call grunt --level=ADVANCED
call grunt --level=ADVANCED --addon=sdkjs-forms --addon=sdkjs-ooxml  --desktop=true

pause