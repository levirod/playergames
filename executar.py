import subprocess
import pandas as pd
import locale
import openpyxl

subprocess.run(["python", "removeassis.py"])
subprocess.run(["python", "resolveplanilha.py"])
subprocess.run(["python", "removepalavra.py"])
subprocess.run(["python", "agrupar.py"])
subprocess.run(["python", "maiusculo.py"])
subprocess.run(["python", "modelarparatemplate.py"])


