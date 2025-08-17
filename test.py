import pefile
import os

def get_exe_version_info_pefile(exe_path):
    if not os.path.exists(exe_path):
        return "dev", ""
    pe = pefile.PE(exe_path)
    build_date = ""
    version = "dev"
    if hasattr(pe, 'FileInfo'):
        for fileinfo in pe.FileInfo:
            # fileinfo may be a list (of StringFileInfo/VarFileInfo)
            if isinstance(fileinfo, list):
                for subinfo in fileinfo:
                    if hasattr(subinfo, 'Key') and subinfo.Key == b'StringFileInfo':
                        for st in subinfo.StringTable:
                            entries = st.entries
                            # Decode bytes to str if needed
                            entries = {k.decode() if isinstance(k, bytes) else k:
                                       v.decode() if isinstance(v, bytes) else v
                                       for k, v in entries.items()}
                            version = entries.get('FileVersion', 'dev')
                            build_date = entries.get('BuildDate', '')
            elif hasattr(fileinfo, 'Key') and fileinfo.Key == b'StringFileInfo':
                for st in fileinfo.StringTable:
                    entries = st.entries
                    # Decode bytes to str if needed
                    entries = {k.decode() if isinstance(k, bytes) else k:
                               v.decode() if isinstance(v, bytes) else v
                               for k, v in entries.items()}
                    version = entries.get('FileVersion', 'dev')
                    build_date = entries.get('BuildDate', '')
    return version, build_date

# Example usage:
exe_path = os.path.abspath(os.path.join(os.path.dirname(__file__), 'dist', 'ActivityLogger.exe'))
version, build_date = get_exe_version_info_pefile(exe_path)
print("Version:", version)
print("BuildDate:", build_date)