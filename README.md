# PowerPoint Mac Math VBA 🍏➕

**A robust, crash-free solution for automating mathematical equations in PowerPoint on macOS.**

## 📝 Description

If you develop VBA macros for PowerPoint on macOS (Office 365/2019+), you likely know the pain: the standard API calls to create equations via VBA are fundamentally broken.

Commands like `Shapes.AddOLEObject` or `CommandBars.ExecuteMso("EquationInsert")` often result in:
* **Run-time error '5': Invalid procedure call or argument**
* **"Selection (unknown member) : Invalid request"**
* Focus issues where the macro cannot "see" the active slide.

### The Solution: "Template Cloning"
Instead of fighting the buggy creation API, this macro uses a **Cloning Approach**:
1.  It locates a pre-defined hidden "dummy" equation (template) on your first slide.
2.  Copies it using standard OS clipboard functions.
3.  Pastes it as a new independent object on the current slide.
4.  Injects your formula using **UnicodeMath** syntax.
5.  Re-compiles the equation to "Professional" display mode.

This method bypasses the unstable creation API entirely, ensuring **100% reliability on Mac**.

---

## 🚀 Installation

1.  Open your PowerPoint presentation (or create a new `.pptm` file).
2.  Press `Opt + F11` (or `Fn + Opt + F11`) to open the Visual Basic Editor.
3.  Go to **Insert** -> **Module**.
4.  Copy and paste the code from `FormulaModule.bas` into the module.
5.  Save the file as **PowerPoint Macro-Enabled Presentation (.pptm)**.

---

## ⚠️ CRITICAL SETUP (Do this once!)

For the macro to work, it needs a "donor" object to clone.

1.  Go to **Slide 1**.
2.  Insert a standard Equation manually (**Insert** -> **Equation**). Type anything (e.g., `x`).
3.  Select this equation object.
4.  Go to the **Home** tab -> **Arrange** -> **Selection Pane**.
5.  Find the selected object in the list and **rename it to**: `MathTemplate`.
6.  *(Optional)* You can hide this object (click the eye icon) or move it off-screen.

---

## 🛠 Usage

1.  Open the VBA Editor.
2.  Locate the `mathText` variable in the code:
    ```vba
    mathText = "x = (-b \pm \sqrt(b^2 - 4ac))/(2a)"
    ```
3.  Edit the formula using **UnicodeMath** syntax.
4.  Run the macro `InsertFormula_ViaCopy`.

---

## 📄 License

[MIT](LICENSE)
