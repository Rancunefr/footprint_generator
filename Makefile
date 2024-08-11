all: Generics.pretty

Generics.pretty: Generics.ods
	./create_footprints_2pins.py Generics

clean:
	rm -fr Generics.pretty
	rm -fr __pycache__
