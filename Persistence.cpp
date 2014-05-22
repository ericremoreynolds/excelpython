//const int NONE = 0;
//const int DICTIONARY = 1;
//const int STRING = 2;
//const int INT32 = 3;
//const int INT64 = 4;
//const int FLOAT64 = 5;
//const int BOOL8 = 6;
//const int LIST = 7;
//const int TUPLE = 8;
//const int DATETIME = 9;
//
//		static PyObject* ReadObject(BinaryReader r)
//		{
//			int tag = r.ReadInt32();
//			switch(tag)
//			{
//			case NONE:
//				return null;
//			case BOOL8:
//				return 0 != r.ReadByte();
//			case INT32:
//				return r.ReadInt32();
//			case INT64:
//				return r.ReadInt64();
//			case FLOAT64:
//				return r.ReadDouble();
//			case STRING:
//				return ReadString(r);
//			case DATETIME:
//				return new DateTime(r.ReadInt64(), DateTimeKind.Utc);
//			case LIST:
//				{
//					int n = r.ReadInt32();
//					List<object> L = new List<object>();
//					for(int i = 0; i < n; i++)
//						L.Add(ReadObject(r));
//					return L;
//				}
//			case TUPLE:
//				{
//					int n = r.ReadInt32();
//					object[] t = new object[n];
//					for(int i = 0; i < n; i++)
//						t[i] = ReadObject(r);
//					return t;
//				}
//			case DICTIONARY:
//				{
//					Dictionary<object, object> d = new Dictionary<object, object>();
//					int n = r.ReadInt32();
//					for(int i = 0; i < n; i++)
//					{
//						string k = (string) ReadObject(r);
//						object v = ReadObject(r);
//						d[k] = v;
//					}
//					return d;
//				}
//			default:
//				throw new Exception("Unexpected tag while reading data.");
//			}
//		}
//
//		static void WriteString(BinaryWriter w, string s)
//		{
//			w.Write(s.Length);
//			w.Write(Encoding.UTF8.GetBytes(s));
//		}
//
//		static void WriteObject(BinaryWriter w, object obj)
//		{
//			if(obj == null)
//			{
//				w.Write(NONE);
//			}
//			else if(obj is double)
//			{
//				w.Write(FLOAT64);
//				w.Write((double) obj);
//			}
//			else if(obj is int)
//			{
//				w.Write(INT32);
//				w.Write((int) obj);
//			}
//			else if(obj is long)
//			{
//				w.Write(INT64);
//				w.Write((long) obj);
//			}
//			else if(obj is string)
//			{
//				string s = (string) obj;
//				w.Write(STRING);
//				WriteString(w, s);
//			}
//			else if(obj is bool)
//			{
//				bool b = (bool) obj;
//				w.Write(BOOL8);
//				w.Write(b ? (byte) 1 : (byte) 0);
//			}
//			else if(obj is DateTime)
//			{
//				DateTime dt = (DateTime) obj;
//				w.Write(DATETIME);
//				w.Write(dt.Ticks);
//			}
//			else if(obj is List<object>)
//			{
//				List<object> list = (List<object>) obj;
//				w.Write(LIST);
//				w.Write(list.Count);
//				foreach(object o in list)
//					WriteObject(w, o);
//			}
//			else if(obj is Dictionary<object, object>)
//			{
//				Dictionary<object, object> dict = (Dictionary<object, object>) obj;
//				w.Write(DICTIONARY);
//				w.Write(dict.Count);
//				foreach(var p in dict)
//				{
//					WriteObject(w, p.Key);
//					WriteObject(w, p.Value);
//				}
//			}
//			else if(obj is Array)
//			{
//				Array arr = (Array) obj;
//				w.Write(TUPLE);
//				w.Write(arr.Length);
//				for(int k = 0; k < arr.Length; k++)
//					WriteObject(w, arr.GetValue(k));
//			}
//			else
//				throw new Exception("Cannot serialize object.");
//		}
//	};