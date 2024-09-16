import cx_Freeze

executables = [cx_Freeze.Executable(
    script = "main.py",
    )]

cx_Freeze.setup(
    name="Ana",
    options={"build_exe": {"packages": ["tkinter","pyautogui", "pandas", "numpy","fsspec"]}},
    executables=executables
)

#python setup.py build