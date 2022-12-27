# Wechat loop send tool

1. Notice about the win32com library, if use Dispatch function with multithread,
you will get an error with com `pywintypes.com_error: (-2147221008, '尚未调用 CoInitialize。', None, None)`
2. Next when you want calling the Dispatch function, before please
    ```Python
        from pythoncom import CoInitialize
        #Other code...
        CoInitialize()
        #COM library calling code...
    ```
3. When multithread running, upon error never show again.
4. Using this script you have to drag the chat frame as independence.
5. Please notice sand time, how do you set it, if send time less than
now time, the send progress will not active at future.
6. If you want see the log of application, you need repackage the script.
7. ```Bash
    pyinstaller -F wechat_loop.pyw  #this is nologging mode
    pyinstaller -F wechat_loop.py   #this is logging mode
   ```
8. Now application version was GM 1.0.1