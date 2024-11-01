# RustClr (Tests)

This project includes tests covering different usage scenarios of the `RustClr` library, which allows loading and executing .NET assemblies in Rust.

## Test Structure

The tests are divided into four main cases:

1. **`test_create_domain`**:
    - Loads a .NET file and creates a custom application domain.
    - Example file: `"file"`
    - Demonstrates the ability to configure a custom application domain when executing a .NET assembly.

2. **`test_with_args`**:
    - Loads a .NET file and executes it with provided string arguments.
    - Example file: `"file"`, arguments: `["test"]`
    - Tests passing arguments to the .NET assembly.

3. **`test_with_runtime`**:
    - Loads a .NET file and specifies a specific runtime version (e.g., `.NET v4`).
    - Example file: `"file"`
    - Tests setting the .NET runtime version during execution.

4. **`test_without_args`**:
    - Loads and runs a .NET file without any additional arguments.
    - Example file: `"file"`
    - Tests basic execution of a .NET assembly without parameters.

## Dependencies

To run the tests, you'll need the following dependencies:

- **`.NET files`**: Ensure that the `.NET` files mentioned in the test cases (e.g., `"file"`) are available in the expected location.

## Running the Tests

To run the tests, use the following command:
```bash
cargo test <test-name> -- --nocapture
```