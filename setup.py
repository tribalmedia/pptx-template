from setuptools import setup
import io
import re

with io.open("pptx_template/__init__.py") as ver:
    metadata = dict(re.findall("__([a-z]+)__\s*=\s*'([^']+)'", ver.read()))

with io.open('README.rst', encoding='ascii') as fp:
    long_description = fp.read()

setup(name='pptx-template',
      version=metadata['version'],
      description='The PowerPoint presentation builder using template.pptx and data(json and csv)',
      long_description=long_description,
      url='http://github.com/m3dev/pptx-template',
      author='Reki Murakami',
      author_email='reki2000@gmail.com',
      license='Apache-2.0',
      packages=['pptx_template'],
      test_suite='test',
      install_requires=['python-pptx==0.6.17', 'pandas >=0.22,<=0.24.0', 'openpyxl>=2.4.7'],
      keywords=['powerpoint', 'ppt', 'pptx'],
      entry_points={ "console_scripts": [ "pptx_template=pptx_template.cli:main"]},
      classifiers=[
        "Development Status :: 3 - Alpha",
        "Topic :: Utilities",
        "Programming Language :: Python :: 2",
        "Programming Language :: Python :: 2.7",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.5",
        "License :: OSI Approved :: Apache Software License",
        "Operating System :: OS Independent"
     ]
)
