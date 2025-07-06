# game_sm
Сила Мысли Игра

Must be run with Python 3.9 or higher.
# Installation

```bash
pip install -r requirements.txt
```
# The new requirements.txt file creation
Attention! This will overwrite the existing "requirements.txt" file.
To silence this warning use ```--strip-extras``` to opt into the new default
or use ```--no-strip-extras``` to retain the existing behavior

```bash
pip install pip-tools
pip-compile requirements.in
```
# The new requirements.in file creation
Collect the main packages used in the project.
It will not include the dependencies of the packages and print them to the console.
Copy the output to the "requirements.in"
```bash
pip install pipdeptree
pipdeptree --warn silence --freeze | findstr /R "^[a-zA-Z0-9]"
```

# The new "requirements.txt" file creation (simple way)
It will includes all the pakages with dependencies of the packages as well.
```bash
pip freeze > requirements.txt
```

