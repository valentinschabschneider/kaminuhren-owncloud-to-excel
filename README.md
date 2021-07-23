# Kaminuhren ownCloud to excel

Converts an ownCloud file / folder structure to an excel file.

## Folder structure

Folder naming scheme is "Kaminuhr" and a four digit number: `Kaminuhr-XXXX`

### Example

```
> root
	...
	Kaminuhr-0173
	Kaminuhr-0174
	Kaminuhr-0175
	...
```

The inner folder stucture contains at least three Files.

- Beschreibung.txt (description)
- \<something>.jpg (image of the clock)
- \<folder-name>.png (qr code)

### Example

```
> Kaminuhr-0174
	Beschreibung.txt
	IMG_20210720_133231__01.jpg
	Kaminuhr-0173.png
	...
```

## Build Executeable

`pipenv install --dev`

`pipenv run python .\setup.py py2exe`
