def xlfunc(f = None, **kwargs):
	def inner(f):
		if f is None:
			f = self
			self = None
		xlfunc = f.__xlfunc__ = {}
		xlfunc["name"] = f.__name__
		xlfunc["doc"] = f.__doc__ if f.__doc__ is not None else "Python function '" + f.__name__ + "' defined in module '" + f.__module__ + "'."
		xlargs = xlfunc["args"] = []
		for vname in f.__code__.co_varnames:
			xlargs.append({
				"name": vname,
				"marshal": "value"
			})
		return f
	import inspect
	if f is not None and len(kwargs) == 0:
		return inner(f)
	else:
		return inner