import sys
import pythoncom
import pywintypes
import win32com.client
import win32com.server.util as serverutil
import win32com.server.dispatcher
import win32com.server.policy
import win32api
import winerror

# --- XLPython object class id ---

clsid = pywintypes.IID(sys.argv[1])

# --- the XLPython class itself ---

class XLPythonOption(object):
	def __init__(self, option, value):
		self.option = option
		self.value = value

class XLPythonObject(object):
	_public_methods_ = []
	def __init__(self, obj):
		self.obj = obj

def FromVariant(var):
	try:
		return win32com.server.util.unwrap(var).obj
	except:
		return var

def ToVariant(obj):
	return win32com.server.util.wrap(XLPythonObject(obj))

class XLPython(object):
	_public_methods_ = [ 'Module', 'Tuple', 'Dict', 'List', 'Obj', 'Str', 'Var', 'Call', 'GetItem', 'SetItem', 'GetAttr', 'SetAttr', 'Eval', 'Exec', 'ShowConsole', 'Reload', 'AddPath' ]
	
	def ShowConsole(self):
		import ctypes
		import sys
		ctypes.windll.kernel32.AllocConsole()
		sys.stdout = open("CONOUT$", "a", 0)
		sys.stderr = open("CONOUT$", "a", 0)
	
	def Module(self, module, *opts):
		options = { 'addpath': None, 'reload': True }
		for opt in opts:
			opt = FromVariant(opt)
			if type(opt) == XLPythonOption:
				print opt.option, opt.value
				options[opt.option] = opt.value
		syspath = list(sys.path) # save the path
		try:
			locals = {}
			if options['addpath'] is not None:
				for i, p in enumerate(options['addpath']):
					sys.path.insert(i, str(p))
			exec "import " + module in locals
			m = locals[module]
			if reload:
				m = __builtins__.reload(m)
		finally:
			# restore the system path
			sys.path = syspath
		return ToVariant(m)
		
	def Reload(self, reload=True):
		return ToVariant(XLPythonOption('reload', reload))

	def AddPath(self, path):
		if isinstance(path, unicode):
			path = path.split(";")
		return ToVariant(XLPythonOption('addpath', path))
		
	def Tuple(self, *elements):
		return ToVariant(tuple((FromVariant(e) for e in elements)))
		
	def Dict(self, *kvpairs):
		if len(kvpairs) % 2 != 0:
			raise Exception("Arguments must be alternating keys and values.")
		n = len(kvpairs) / 2
		d = {}
		for k in range(n):
			key = FromVariant(kvpairs[2*k])
			value = FromVariant(kvpairs[2*k+1])
			d[key] = value
		return ToVariant(d)
		
	def List(self, *elements):
		return ToVariant(list((FromVariant(e) for e in elements)))
		
	def Obj(self, var):
		return ToVariant(FromVariant(var))
		
	def Str(self, obj):
		return str(FromVariant(obj))
		
	def Var(self, obj):
		return FromVariant(obj)
		
	def Call(self, obj, *args):
		obj = FromVariant(obj)
		method = None
		pargs = ()
		kwargs = {}
		for arg in args:
			arg = FromVariant(arg)
			if isinstance(arg, unicode):
				method = arg
			elif isinstance(arg, tuple):
				pargs = arg
			elif isinstance(arg, dict):
				kwargs = arg
		if method is None:
			return ToVariant(obj(*pargs, **kwargs))
		else:
			return ToVariant(getattr(obj, method)(*pargs, **kwargs))
			
	def GetItem(self, obj, key):
		obj = FromVariant(obj)
		key = FromVariant(key)
		return ToVariant(obj[key])
		
	def SetItem(self, obj, key, value):
		obj = FromVariant(obj)
		key = FromVariant(key)
		value = FromVariant(value)
		obj[key] = value
		
	def GetAttr(self, obj, attr):
		obj = FromVariant(obj)
		attr = FromVariant(attr)
		return ToVariant(getattr(obj, attr))
		
	def SetAttr(self, obj, attr, value):
		obj = FromVariant(obj)
		attr = FromVariant(attr)
		value = FromVariant(value)
		setattr(obj, attr, value)
		
	def Eval(self, expr, *args):
		globals = None
		locals = None
		for arg in args:
			arg = FromVariant(arg)
			if type(arg) is dict:
				if globals is None:
					globals = arg
				elif locals is None:
					locals = arg
				else:
					raise Exception("Eval can be called with at most 2 dictionary arguments")
			else:
				pass
		return ToVariant(eval(expr, globals, locals))
		
	def Exec(self, stmt, *args):
		globals = None
		locals = None	
		for arg in args:
			arg = FromVariant(arg)
			if type(arg) is dict:
				if globals is None:
					globals = arg
				elif locals is None:
					locals = arg
				else:
					raise Exception("Exec can be called with at most 2 dictionary arguments")
			else:
				pass
		exec stmt in globals, locals
		
# --- ovveride CreateInstance in default policy to instantiate the XLPython object ---

BaseDefaultPolicy = win32com.server.policy.DefaultPolicy

class MyPolicy(BaseDefaultPolicy):
	def _CreateInstance_(self, reqClsid, reqIID):
		if reqClsid == clsid:
			return serverutil.wrap(XLPython(), reqIID)
			print reqClsid
		else:
			return BaseDefaultPolicy._CreateInstance_(self, clsid, reqIID)

win32com.server.policy.DefaultPolicy = MyPolicy

# --- create the class factory and register it

factory = pythoncom.MakePyFactory(clsid)

clsctx = pythoncom.CLSCTX_LOCAL_SERVER
flags = pythoncom.REGCLS_MULTIPLEUSE | pythoncom.REGCLS_SUSPENDED
revokeId = pythoncom.CoRegisterClassObject(clsid, factory, clsctx, flags)

pythoncom.EnableQuitMessage(win32api.GetCurrentThreadId())	
pythoncom.CoResumeClassObjects()

pythoncom.PumpMessages()

pythoncom.CoRevokeClassObject(revokeId)
pythoncom.CoUninitialize()
