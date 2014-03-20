from setuptools import setup

setup(name='sugar2xls',
      version='0.2.5',
      scripts=['sugar2xls.py'],
      description='Helper tool dedicated to the retrieval of Sugar CRM data for reporting purpose',
      author='David Loureiro',
      author_email='david.loureiro1@gmail.com',
      url='https://github.com/dloureiro/sugar2xls',
      dependency_links=['http://github.com/luisbarrueco/python_webservices_library/tarball/master#egg=sugarcrm-0.2.0'],
      install_requires=["xlwt>=0.7.5","sugarcrm>=0.2.0"],
      classifiers=[
      'Development Status :: 3 - Alpha',
      'Environment :: Console',
      'License :: OSI Approved :: GNU General Public License v3',
      'Natural Language :: English',
      'Operating System :: MacOS :: MacOS X',
      'Operating System :: Microsoft :: Windows',
      'Operating System :: POSIX',
      'Programming Language :: Python',
      'Topic :: Documentation',
      'Topic :: Internet :: WWW/HTTP',
      'Topic :: Multimedia :: Graphics',
      'Topic :: Software Development :: Documentation',
      'Topic :: Text Editors :: Documentation',
      'Topic :: Text Editors :: Text Processing',
      'Topic :: Utilities']
      )