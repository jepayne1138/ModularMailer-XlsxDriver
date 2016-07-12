from setuptools import setup

setup(
    name='ModularMailer-XlsxDriver',
    packages=['xlsx_driver'],
    version='0.0a1',
    description='ModularMailer plugin for interacting with xlsx files',
    author='James Payne',
    author_email='jepayne1138@gmail.com',
    url='https://github.com/jepayne1138/ModularMailer-XlsxDriver',
    license='BSD-new',
    download_url='https://github.com/jepayne1138/ModularMailer-XlsxDriver/tarball/0.0a1',
    keywords='plugin xlsx',
    install_requires=['openpyxl', 'ModularMailer'],
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Programming Language :: Python',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: Implementation :: CPython',
        'License :: OSI Approved :: BSD License',
        'Intended Audience :: End Users/Desktop',
        'Environment :: Plugins',
    ],
)
