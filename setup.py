#/usr/bin/env python3.5
# vim: set fileencoding=utf-8 fileformat=unix :

from setuptools import setup

from exceltable import \
        __author__, __copyright__, __license__, __version__, __email__


setup(
    name = "exceltable",
    version = __version__,
    author = __author__,
    author_email = __email__,
    license = __license__,
    platforms = ["generic"],
    packages = ["exceltable"],
    install_requires = ["xlrd>=0.9.4", "docopt>=0.6.2"],
    entry_points = dict(
            console_scripts = ["exceltable=exceltable.command:main"],
            ),
    )
