import sys
import os

def get_platform():
    """
    Detects and returns the current operating system platform.
    Returns: A string indicating 'Windows', 'macOS', or 'Other'
    """
    if sys.platform.startswith('win32') or sys.platform.startswith('cygwin'):
        return 'Windows'
    elif sys.platform.startswith('darwin'):
        return 'macOS'
    else:
        return 'Other'

def get_platform_details():
    """
    Returns detailed information about the current platform.
    """
    return {
        'platform': get_platform(),
        'system': os.name,
        'platform_system': sys.platform,
        'python_version': sys.version,
        'path_separator': os.sep,
        'line_separator': os.linesep
    }

def get_chrome_profile_path():
    """
    Returns the platform-specific Chrome profile path.
    """
    if get_platform() == 'Windows':
        return os.path.join(os.path.expanduser("~/Desktop"), "Chrome Profile 1")
    elif get_platform() == 'macOS':
        return os.path.join(os.path.expanduser("~/Desktop"), "Chrome Profile (PlayAI API)")
    else:
        return os.path.join(os.path.expanduser("~/Desktop"), "Chrome Profile 1")

def get_celebrity_vo_path():
    """
    Returns the platform-specific path for Celebrity Voice Overs.
    """
    if get_platform() == 'Windows':
        return "E:\\Celebrity Voice Overs"
    elif get_platform() == 'macOS':
        return os.path.join(os.path.expanduser("~/Desktop"), "Celebrity Voice Overs")
    else:
        return os.path.join(os.path.expanduser("~/Desktop"), "Celebrity Voice Overs")

if __name__ == '__main__':
    # Simple usage example
    print(f"Current platform: {get_platform()}")
    print("\nDetailed platform information:")
    for key, value in get_platform_details().items():
        print(f"{key}: {value}")
    print(f"\nChrome Profile Path: {get_chrome_profile_path()}")
    print(f"Celebrity Voice Overs Path: {get_celebrity_vo_path()}") 