.PHONY: build install verify

build:
	./scripts/build_skill.sh

install: build
	./scripts/install_skill.sh

verify:
	./scripts/verify_skill.sh
