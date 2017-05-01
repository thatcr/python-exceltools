"""
install exceltools, a package for deploying python excel extensions
"""
from setuptools import setup, Extension, find_packages

setup(
    name='python-exceltools',
    use_vcs_version={'increment': '0.1'},
    author='Robert Thatcher',
    author_email='r.thatcher@cf-partners.com',
    ext_modules=[
        Extension('exceltools._addin', ['exceltools/_addin.c']),
    ],
    packages=find_packages(),
    package_data={
        '': ["*.xll"]
    },
    entry_points={
        'console_scripts': [
            'excel_entry_points = exceltools.entrypoints:write_entry_points'
        ]
    },
    license='LICENSE.txt',
    description='Install Python Excel addins.',
    test_suite="nose.collector",
    install_requires=[
        'setuptools>=7.0'
    ],
    setup_requires=['hgtools'],
    extras_require={}
)
