# Contributing to ecuacion-utils

Thank you for your interest in contributing!

## Reporting Bugs

Please open a [GitHub Issue](https://github.com/ecuacion-jp/ecuacion-utils/issues) with:

- A clear description of the problem
- Steps to reproduce
- Expected vs. actual behavior
- Java version and library version

## Suggesting Features

Open a [GitHub Issue](https://github.com/ecuacion-jp/ecuacion-utils/issues) describing the use case and the proposed behavior. Feel free to discuss before submitting a pull request.

## Submitting Pull Requests

1. Fork the repository and create a branch from `main`.
2. Make sure the build passes: `mvn clean verify`
3. Add or update tests for any changed behavior.
4. Open a pull request with a clear description of what was changed and why.

## Build Requirements

- JDK 21 or above
- Maven 3.9 or above
- `ecuacion-lib` must be checked out as a sibling directory and installed locally before building:
  ```
  cd ../ecuacion-lib && mvn install -DskipTests
  ```

## Code Style

The project enforces Checkstyle and SpotBugs rules at build time. Run `mvn clean verify` locally before submitting — the CI will catch violations, but catching them early saves time.

## License

By submitting a pull request, you agree that your contribution will be licensed under the [Apache License 2.0](LICENSE.txt).
