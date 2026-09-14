import os
import sys
import streamlit.web.cli as stcli

def get_base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

if __name__ == '__main__':
    base_path = get_base_path()
    script_path = os.path.join(base_path, 'streamlit_app.py')
    
    sys.argv = [
        "streamlit",
        "run",
        script_path,
        "--global.developmentMode=false"
    ]
    sys.exit(stcli.main())