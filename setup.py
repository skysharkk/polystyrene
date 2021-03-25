from distutils import version
from distutils.util import execute
from cx_Freeze import setup, Executable

setup(
    name="polystyrene",
    version="1.0",
    description="Расскладка ППТ",
    executables=[Executable("polystyrene.py")]
)
