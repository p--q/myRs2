#!/opt/libreoffice5.4/program/python
# -*- coding: utf-8 -*-
# Created by modifying urlimport.py in 10.11.loading_modules_from_a_remote_machine_using_import_hooks of Python Cookbook 3rd Edition.
# 一旦LibreOfficeを終了させないとimportはキャッシュが使われるのでデバッグ時は必ずLibreOfficeを終了すること!!!
# インポートするパッケジーには__init__.pyが必要。
import sys
import importlib.abc
from types import ModuleType
def _get_links(simplefileaccess, url):  # url内のファイル名とフォルダ名のリストを返す関数。
	foldercontents = simplefileaccess.getFolderContents(url, True)  # url内のファイルとフォルダをすべて取得。フルパスで返ってくる。
	tdocpath = "".join((url, "/"))  # 除去するパスの部分を作成。
	return [content.replace(tdocpath, "") for content in foldercontents]  # ファイル名かフォルダ名だけのリストにして返す。
class UrlMetaFinder(importlib.abc.MetaPathFinder):  # meta path finderの実装。
	def __init__(self, simplefileaccess, baseurl):
		self._simplefileaccess = simplefileaccess  # LibreOfficeドキュメント内のファイルにアクセスするためのsimplefileaccess
		self._baseurl = baseurl  # モジュールを探すパス
		self._links   = {}  # baseurl内のファイル名とフォルダ名のリストのキャッシュにする辞書。。
		self._loaders = {baseurl: UrlModuleLoader(simplefileaccess, baseurl)}  # ローダーのキャッシュにする辞書。
	def find_module(self, fullname, path=None):  # find_moduleはPython3.4で撤廃だが、find_spec()にしてもそのままではうまく動かない。
		if path is None:
			baseurl = self._baseurl
		else:
			if not path[0].startswith(self._baseurl):
				return None
			baseurl = path[0]
		parts = fullname.split('.')
		basename = parts[-1]
		if basename not in self._links:  # Check link cache
			self._links[baseurl] = _get_links(self._simplefileaccess, baseurl)
		if basename in self._links[baseurl]:  # Check if it's a package。パッケージの時。
			fullurl = "/".join((self._baseurl, basename))
			loader = UrlPackageLoader(self._simplefileaccess, fullurl)
			try:  # Attempt to load the package (which accesses __init__.py)
				loader.load_module(fullname)
				self._links[fullurl] = _get_links(self._simplefileaccess, fullurl)
				self._loaders[fullurl] = UrlModuleLoader(self._simplefileaccess, fullurl)
			except ImportError:
				loader = None
			return loader
		filename = "".join((basename, '.py'))
		if filename in self._links[baseurl]:  # A normal module
			return self._loaders[baseurl]
		else:
			return None
	def invalidate_caches(self):
		self._links.clear()
class UrlModuleLoader(importlib.abc.SourceLoader):  # Module Loader for a URL
	def __init__(self, simplefileaccess, baseurl):
		self._simplefileaccess = simplefileaccess  # LibreOfficeドキュメント内のファイルにアクセスするためのsimplefileaccess
		self._baseurl = baseurl  # モジュールを探すパス
		self._source_cache = {}  # ソースのキャッシュの辞書。
	def module_repr(self, module):  # モジュールを表す文字列を返す。
		return '<urlmodule {} from {}>'.format(module.__name__, module.__file__)
	def load_module(self, fullname):  # Required method。引数はimport文で使うフルネーム。
		code = self.get_code(fullname)  # モジュールのコードオブジェクトを取得。
		mod = sys.modules.setdefault(fullname, ModuleType(fullname))  # 辞書sys.modulesにキーfullnameなければ値を代入して値を取得。
		mod.__file__ = self.get_filename(fullname)  # ソースファイルへのフルパスを取得。
		mod.__loader__ = self  # ローダーを取得。
		mod.__package__ = fullname.rpartition('.')[0]  # パッケージ名を取得。.区切りがないときは空文字が入る。
		exec(code, mod.__dict__)  # コードオブジェクトを実行する。
		return mod  # モジュールオブジェクトを返す。
	def get_code(self, fullname):  # モジュールのコードオブジェクトを返す。Optional extensions。引数はimport文で使うフルネーム。
		src = self.get_source(fullname)
		return compile(src, self.get_filename(fullname), 'exec')
	def get_data(self, path):  # バイナリ文字列を返す。
		pass
	def get_filename(self, fullname):  # ソースファイルへのフルパスを返す。引数はimport文で使うフルネーム。
		return "".join((self._baseurl, '/', fullname.split('.')[-1], '.py'))
	def get_source(self, fullname):  # モジュールのソースをテキストで返す。
		filename = self.get_filename(fullname)  # ソースファイルへのフルパス。
		if filename in self._source_cache:  # すでにキャッシュがあればそれを返して終わる。
			return self._source_cache[filename]
		try:
			inputstream = self._simplefileaccess.openFileRead(filename)  # ソースファイルのインプットストリームを取得。
			dummy, b = inputstream.readBytes([], inputstream.available())  # simplefileaccess.getSize(module_tdocurl)は0が返る。
			source = bytes(b).decode("utf-8")  # モジュールのソースファイルをutf-8のテキストで取得。
			self._source_cache[filename] = source  # ソースをキャッシュに取得。
			return source  # ソースのテキストを返す。
		except:
			raise ImportError("Can't load {}".format(filename))
	def is_package(self, fullname):  # パッケージの時はTrueを返す。
		return False
class UrlPackageLoader(UrlModuleLoader):  # Package loader for a URL
	def load_module(self, fullname):  # fullnameはパッケージのときはフォルダ名に該当する。
		mod = super().load_module(fullname)  # __init__.pyを実行する。
		mod.__path__ = [self._baseurl]  # パッケージ内の検索パスを指定する文字列のリスト
		mod.__package__ = fullname  # フォルダ名を入れる。
	def get_filename(self, fullname):  # パッケージの__init__.pyを返す。
		return "/".join((self._baseurl, '__init__.py'))
	def is_package(self, fullname):  # パッケージの時はTrueを返す。
		return True
_installed_meta_cache = {}  # meta path finderを入れておくグローバル辞書。重複を防ぐ目的。
def install_meta(simplefileaccess, address):  # Utility functions for installing the loader
	if address not in _installed_meta_cache:  # グローバル辞書にないパスの時。
		finder = UrlMetaFinder(simplefileaccess, address)  # meta path finder。モジュールを探すクラスをインスタンス化。
		_installed_meta_cache[address] = finder  # グローバル辞書にmeta path finderを登録。
		sys.meta_path.append(finder)  # meta path finderをsys.meta_pathに登録。
def remove_meta(address):  # Utility functions for uninstalling the loader
	if address in _installed_meta_cache:
		finder = _installed_meta_cache.pop(address)
		sys.meta_path.remove(finder)
