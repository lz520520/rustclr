use rustclr::{
    RustClrEnv, ClrOutput, 
    Invocation, Variant,
};

fn sub() -> Result<(), Box<dyn std::error::Error>> {
    // Create and initialize the CLR environment
    let clr = RustClrEnv::new(None)?;
    let mscorlib = clr.app_domain.load_lib("mscorlib")?;
    let console = mscorlib.resolve_type("System.Console")?;

    // Set up output redirection
    let mut clr_output = ClrOutput::new(&mscorlib);
    clr_output.redirect()?;

    // Prepare the arguments
    let args = vec!["Hello World".to_variant()];

    // Invoke the WriteLine method
    console.invoke("WriteLine", None, Some(args), Invocation::Static)?;
    console.invoke("WriteLine", None, Some(vec!["Hello World111".to_variant()]), Invocation::Static)?;


    // Restore the original output and capture redirected content
    clr_output.restore()?;
    let output = clr_output.capture()?;

    print!("output: {output}");

    Ok(())
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    for _i in 0..2 {
        sub()?;
    }

    Ok(())
}