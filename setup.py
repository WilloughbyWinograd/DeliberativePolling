import setuptools
  
with open("README.md", "r") as fh:
    description = fh.read()
  
setuptools.setup(
    name="DeliberativePolling",
    version="0.1.9",
    author="The Deliberative Democracy Lab at Stanford University",
    author_email="deliberation@stanford.edu",
    packages=["DeliberativePolling"],
    description="A package for comparing survey data between samples from Deliberative Polling experiments.",
    long_description=description,
    long_description_content_type="text/markdown",
    url="https://github.com/WilloughbyWinograd/DeliberativePolling",
    license='MIT',
    python_requires='>=3.8',
    install_requires=[
        "pandas",
        "numpy",
        "openpyxl",
        "pyreadstat",
        "statsmodels",
        "scipy",
        "docx",
        "tqdm",
    ]
)