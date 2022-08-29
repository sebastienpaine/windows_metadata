from setuptools import setup, find_packages


setup(
    name='windows_metadata',
    version='0.1.0',
    license='MIT',
    author="Sébastien Paine",
    author_email='sebastienpaine@me.com',
    packages=find_packages('src'),
    package_dir={'': 'src'},
    url='https://github.com/gmyrianthous/example-publish-pypi',
    keywords='windows metadata attributes details',
    install_requires=[
          'pywin32',
      ],

)