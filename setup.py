from setuptools import setup, find_packages

setup(
    name = "django-report-utils",
    version = "0.3.9",
    author = "David Burke",
    author_email = "david@burkesoftware.com",
    description = ("Common report functions used by django-report-builder and django-report-scaffold."),
    license = "BSD",
    keywords = "django report",
    url = "https://github.com/burke-software/django-report-utils",
    packages=find_packages(),
    include_package_data=True,
    test_suite='setuptest.setuptest.SetupTestSuite',
    tests_require=(
        'django-setuptest',
        'south',
    ),
    classifiers=[
        'Environment :: Web Environment',
        'Framework :: Django',
        'Programming Language :: Python',
        'Intended Audience :: Developers',
        'Intended Audience :: System Administrators',
        "License :: OSI Approved :: BSD License",
    ],
    install_requires=[
        'django',
        'six',
    ]
)
