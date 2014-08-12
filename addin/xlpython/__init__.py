def xlfunc(f):
	if not hasattr(f, "__xlfunc__"):
		xlf = f.__xlfunc__ = {}
		xlf = f.__xlfunc__ = {}
		xlf["name"] = f.__name__
		xlf["doc"] = f.__doc__ if f.__doc__ is not None else "Python function '" + f.__name__ + "' defined in module '" + f.__module__ + "'."
		xlargs = xlf["args"] = []
		xlargmap = xlf["argmap"] = {}
		for vpos, vname in enumerate(f.__code__.co_varnames[:f.__code__.co_argcount]):
			xlargs.append({
				"name": vname,
				"pos": vpos,
				"marshal": "var",
				"range": False,
				"dtype": None,
				"dims": -1
			})
			xlargmap[vname] = xlargs[-1]
		xlf["ret"] = {
			"marshal": "auto"
		}
	return f
		
def xlret(marshal):
	def inner(f):
		xlf = xlfunc(f).__xlfunc__
		xlf["ret"]["marshal"] = marshal
		return f
	return inner
	
def xlarg(arg, marshal="var", dims=-1, dtype=None, range=False):
	def inner(f):
		xlf = xlfunc(f).__xlfunc__
		xlarg = xlf["argmap"][arg]
		xlarg["marshal"] = marshal
		xlarg["dtype"] = dtype
		xlarg["dims"] = dims
		xlarg["range"] = range
		return f
	return inner