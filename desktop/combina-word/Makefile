.PHONY: install run test deb clean

install:
	python3 -m pip install -r requirements.txt

run:
	python3 main.py

test:
	python3 -m pytest -q

deb:
	chmod +x packaging/build-deb.sh && packaging/build-deb.sh

clean:
	rm -rf build dist .venv-build *.egg-info .pytest_cache
