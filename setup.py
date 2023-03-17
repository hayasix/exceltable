#/usr/bin/env python3
# vim: set fileencoding=utf-8 fileformat=unix expandtab :

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
    python_requires=">=3.7",
    packages = ["exceltable"],
    install_requires = ["openpyxl>=3.0.0", "docopt>=0.6.2", "msoffcrypto-tool>=5.0.0"],
    entry_points = dict(
            console_scripts = ["exceltable=exceltable.command:__main__"],
            ),
    )
