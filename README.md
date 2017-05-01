*** Important

In order to obey setup_requires, your ~/pydistutils.cfg must specify the 'simple' form of the pypy url, as
setup_require is run by distutils, not by pip.

It's also important not to have only wheel files in your local pypi for those packages, as setuptools doesn't know
what to do with them.

```
[easy_install]
index_url = http://pypi/root/prod/+simple/
```