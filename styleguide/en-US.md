# Excel Formula Style Guide

## 1. Introduction

### 1.1 Style Guide Version

1.0 draft

### 1.2 Terminology notes

In this Document, the following **BOLD AND CAPITALIZED** key words are Requirement Levels defined in [RFC 2119](https://datatracker.ietf.org/doc/html/rfc2119).

- **MUST**, **REQUIRED**, **SHALL**
- **MUST NOT**, **SHALL NOT**
- **SHOULD**, **RECOMMENDED**
- **SHOULD NOT**, **NOT RECOMMENDED**
- **MAY**, **OPTIONAL**

## 2. Editor

Advanced Formula Environment(AFE) **MUST** be used.

Reason: This style guide is optimized for using AFE.

Remark: AFE is currently in [Excel Labs](https://github.com/microsoft/Excel-Labs/).

## 3. Formatting

### 3.1 Block indentation

4 space(U+0020) **MUST** be used.

Reason: This is AFE's default.

### 3.2 Column limit

110 characters based on half-width characters are **RECOMMENDED**.

Reason: On a Full-HD monitor and 100% scale, approximately 114 half-width characters are displayed in the maximized AFE Names and Modules editor.

Exceptions: When line wrapping compromise readability.

### 3.3 Comma style

Trailing style **MUST** be used.

Reason: Inside the `AND`, `OR`, `IFS`, `LET`, `SWITCH` and other functions, the beginning of the expression is aligned.

```excel
// Good
=AND(
    boolFormula1,
    boolFormula2,
    boolFormula3
)

// Bad | Leading style
=AND(
    boolFormula1 // Not aligned.
    , boolFormula2
    , boolFormula3
)
```

### 3.4 Empty line

#### 3.4.1 Inside formula

**NOT RECOMMENDED**.

Reason: The positional relationship of lines should have no meaning because Excel formulas aren't procedural language.

```excel
// Good
=LET(
    variable1, "value1",
    variable2, "value2",
    variable3, "value3",
    "calculation"
)

// Bad
=LET(
    variable1, "value1",
    variable2, "value2",

    variable3, "value3", // Delete the line avobe.
    "calculation"
)
```

#### 3.4.2 Between named object in a module

##### ranges, constants, formulas

1 or more empty line **MAY** be inserted.

Reason: Because declarations of named objects are procedural, the positional relationship of declarations may have some meaning.

```excel
// Good
EXAMPLEMODULE.MODULENAME = "EXAMPLEMODULE";
EXAMPLEMODULE.MODULEVERSION = "1.0.0";

EXAMPLEMODULE.TAXRATE = 0.1;
```

##### functions

1 or more empty line **MUST** be inserted.

Reason: Definitions of function are often long.

```excel
// Good
EXAMPLEMODULE.EXAMPLEFUNCTION1 = LAMBDA(
    "calculation"
);

EXAMPLEMODULE.EXAMPLEFUNCTION2 = LAMBDA(
    "calculation"
);
```

```excel
// Bad
EXAMPLEMODULE.EXAMPLEFUNCTION1 = LAMBDA(
    "calculation"
);
EXAMPLEMODULE.EXAMPLEFUNCTION2 = LAMBDA( // Insert empty line here.
    "calculation"
);
```

### 3.5 Line wrapping

Breaking line rule is as follows:

1. A line break **MUST** be after the operator, and an indent **MUST** be added before continuation of the formula.

    ```excel
    // Good
    =LAMBDA(
        someLongForluma1 +
            someLongForluma2 -
            someLongForluma3
    )

    // Bad | The break is before the operator.
    =LAMBDA(
        someLongForluma1
            + someLongForluma2
            - someLongForluma3
    )
    
    // Bad | No indents.
    =LAMBDA(
        someLongForluma1 +
        someLongForluma2 -
        someLongForluma3
    )
    ```

2. A left brace `(` and a matching right brace `)` **MUST** be same indentation level, and indentation level between braces **MUST** be deeper than braces.

    ```excel
    // Good
    =AND(
        boolFormula1,
        boolFormula2,
        boolFormula3
    )

    // Bad | The left brace and right brace are in different indentation level.
    =AND(boolFormula1,
        boolFormula2,
        boolFormula3)
    
    // Bad | The indentation level between braces isn't deeper than braces.
    =AND(
    boolFormula1,
    boolFormula2,
    boolFormula3
    )
    ```

### 3.6 Space

1. Trailing spaces **MUST NOT** be present.

    ```excel
    ="some formula"/* Good */
    ="some formula"     /* Bad */
    ```

2. A comma `,` and semicolon `;` as argument list or array separator has spaces with the following rules:
    - If the following argument is present, it **MUST** have 1 space after it but **MUST NOT** before it.

        ```excel
        =LET(
            variableA, "valueA", // Good
            variableB,"valueB", // Bad
            variableC ,"valueC", // Bad
            "calculation"
        )
        ={1, 2, 3; 4, 5, 6; 7, 8, 9} // Good
        ={1,2,3;4,5,6;7,8,9} // Bad
        ={1 ,2 ,3 ;4 ,5 ,6 ;7 ,8 ,9} // Bad
        ```

    - If the following argument is omitted, **MUST NOT** have any spaces between the next comma `,` or brace `)`.

        ```excel
        =INDEX(A1:C3,, 3) // Good
        =INDEX(A1:C3, , 3) // Bad
        =INDEX(A1:C3,2,) // Good
        =INDEX(A1:C3,2, ) // Bad
        ```

3. An operator has spaces with the following rules:
    - The following operators **MUST** have 1 space before and after it.
        - `+` and `-` as binary operators
        - all comparison operators
        - `&`

        ```excel
        =1 + 2 // Good
        =1+2 // Bad
        =1 = 2 // Good
        =1=2 // Bad
        ="a" & "b" // Good
        ="a"&"b" // Bad
        ```

    - The following operators **MUST NOT** any spaces before and after it.
        - `^`

        ```excel
        =1^2 // Good
        =1 ^ 2 // Bad
        ```

    - The following operators **MAY** have 1 space before and after it.
        - `*`, `/`

        ```excel
        =1 * 2 // Good
        =1*2 // Good
        =1 / 2 // Good
        =1/2 // Good
        =1 / 2 // Good
        =1/2 // Good
        ```

4. A Comment has spaces with the following rules:
    - There **MUST** be 1 space before and 1 or more spaces after the end-of-line comment mark `//`.
    - There **MUST** be 1 or more space after the open-block comment mark `/*` and before close-block comment mark `*/`.

    ```excel
    ="some formula" // Good
    ="some formula"//Bad
    ="some formula"/* Good */
    ="some formula"/*Bad*/
    ```

5. An equal `=` as assignment operator **MUST** have 1 space before and after it.

    ```excel
    // Good
    EXAMPLEMODULE.EXAMPLEFUNCTION = LAMBDA(
        "calculation"
    );

    // Bad
    EXAMPLEMODULE.EXAMPLEFUNCTION=LAMBDA(
        "calculation"
    );
    ```

## 4. Naming

### 4.1 Rules common to all identifiers

- Only half-width alphabets, numbers and Low Line `_` **MUST** be used.
- An alphabet **MUST** be useds for the first letter.
- Meaningful name is **RECOMMENDED**.
- Uncommon abbreviations **SHOULD** be avoided.

```excel
// Good
=LET(
    customerName, "John Doe", // Meaningful and no abbreviations.
    customerID, "12345", // "ID" is abbreviation for IDentifier but it's well-known.
    "calculation"
)

// Bad
=LET(
    a, "John Doe", // "a" has no meaning.
    name, "John Doe", // Although some meaningful, only "name" is insufficient and ambiguous.
    cstNm, "John Doe", // Although meaningful, "cst" and "nm" aren't well-known abbreviations.
    "calculation"
)
```

### 4.2 Variable name in `LET` functiuon

Variable names in `LET` function as follows:

#### 4.2.1 Local value(range, contant or formula)

lowerCamelCase **MUST** be used. For example, `customerName`.

Reason: The sample codes in the official [`LET` function reference](https://support.microsoft.com/en-us/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999) does so.

```excel
// Good
=LET(
    lowerCamelCase, "lower camel case",
    "calculation"
)

// Bad
=LET(
    nocapitalized, "no capitalized",
    ALLCAPITALIZED, "all capitalized",
    UpperCamelCase, "upper camel case",
    snake_case, "snake case",
    SCREAMING_SNAKE_CASE, "screaming snake case",
    "calculation"
)
```

#### 4.2.1 Local function

1. snake_case **MUST** be used. For example, `calculate_price`.
2. The name **MUST** end with Low Line `_` if it consists of a single word.

Reason:

1. To help distinguish them from global functions.
2. To avoid naming collision with Excel built-in function. Their names don't contain Low Line `_`.

```excel
// Good
=LET(
    snake_case, LAMBDA("snake case"),
    "calculation"
)

// Bad
=LET(
    nocapitalized, LAMBDA("no capitalized"),
    ALLCAPITALIZED, LAMBDA("all capitalized"),
    lowerCamelCase, LAMBDA("lower camel case"),
    UpperCamelCase, LAMBDA("upper camel case"),
    CALCULATE_PRICE, LAMBDA("screaming snake case"),
    "calculation"
)
```

### 4.3 Argument name in `LAMBDA` functiuon

This **SHALL** be according to the local variable name in `LET` functiuon.

### 4.4 Single named object

A name of named object(range, constant, formula, function) **MUST** have a prefix with an Low Line `_`. For example, prefix `EXAMPLECOMPANY` and base function name `CALCULATEPRICE` into `EXAMPLECOMPANY_CALCULATEPRICE`.

Reason: To avoid naming collision in the future. In particular, all Excel built-in function names don't contain Low Line `_`.

A name of named object **MUST NOT** contain Full Stop `.`.

Reason: Named objects in the same namespace must always be able to be imported together, but single functions can't.

### 4.5 Named object in Workbook module

A name of named object(range, constant, formula, function) **MUST** have a prefix with an Low Line `_`.

Reason: For the same reason as a single named object.

A name of named object **SHOULD NOT** contain Full Stop `.`.

Reason: To maintain future portability. Full Stop `.` is allowed in Workbook module but not in other modules.

Exceptions: When use only 1 Full Stop `.` as namespace separator.

### 4.6 Named object in other modules

Module name and name of named object(range, constant, formula, function) in module **MAY** contain Low Line `_`.

Reason: Excel built-in functions use namespaces, but only a few, so the risk of name collisions is low.

### 5. Function structure

#### 5.1 Entire structure

A function structed as follows:

1. The entire function definition.
2. Arguments(optional).
3. Calculaion part with arguments validation. It **SHOULD** be enclosed in `LET` function.  
    Exceptions: No arguments validation or argument count is 0 or 1.
4. Defining variables to impute validation results and validating arguments.
5. Calculaion part with guard clause for arguments validation. It **SHOULD** be enclosed in `IF` function.  
    Exceptions: No arguments validation.
6. Guard clause for arguments validation. The result of argument validation **MUST** be verified and a `#VALUE!` error or error message returned if invalid.
7. Main Calculaion.

```excel
// Raise error basic style
/* 1        */RAISE_ERROR_BASIC_STYLE = LAMBDA(
/* |2       */    numberArgument,
/* ||       */    optionaltextArgument,
/* |2       */    enumerateTextArgument,
/* | 3      */    LET(
/* | |4     */        argumentIsInvalid, OR(
/* | ||     */            OR(ISOMITTED(numberArgument), NOT(ISNUMBER(numberArgument))), // numberArgument validation
/* | ||     */            AND(NOT(ISOMITTED(optionaltextArgument)), NOT(ISTEXT(optionaltextArgument))), // optionaltextArgument validation
/* | ||     */            ISERROR(LET(list, {"apple", "banana", "citrus"},FILTER(list, list=enumerateTextArgument))) // enumerateTextArgument validation
/* | |4     */        ),
/* | |  5   */        IF(
/* | |  |6  */            argumentIsInvalid, #VALUE!,
/* | |  | 7 */            "calculation"
/* | |  | | */            // calculation
/* | |  | 7 */            // calculation
/* | |  5   */        )
/* | 3      */    )
/* 1        */);

// Return error message basic style
/* 1        */RETURN_ERROR_MESSAGE_BASIC_STYLE = LAMBDA(
/* |2       */    numberArgument,
/* ||       */    optionaltextArgument,
/* |2       */    enumerateTextArgument,
/* | 3      */    LET(
/* | |4     */        errorMessage, IFS(
/* | ||     */            OR(ISOMITTED(numberArgument), NOT(ISNUMBER(numberArgument))), "numberArgument is invalid.", // numberArgument validation
/* | ||     */            AND(NOT(ISOMITTED(optionaltextArgument)), NOT(ISTEXT(optionaltextArgument))), "optionaltextArgument is invalid.", // optionaltextArgument validation
/* | ||     */            ISERROR(LET(list, {"apple", "banana", "citrus"},FILTER(list, list=enumerateTextArgument))), "enumerateTextArgument is invalid.", // enumerateTextArgument validation
/* | ||     */            TRUE, ""
/* | |4     */        ),
/* | |  5   */        IF(
/* | |  |6  */            errorMessage <> "", errorMessage,
/* | |  | 7 */            "calculation"
/* | |  | | */            // calculation
/* | |  | 7 */            // calculation
/* | |  5   */        )
/* | 3      */    )
/* 1        */);

// Raise error one validation style
/* 1        */RAISE_ERROR_ONE_VALIDATION_STYLE = LAMBDA(
/* |2       */    numberArgument,
/* |    5   */    IF(
/* |    |6  */        NOT(ISNUMBER(numberArgument))/* numberArgument validation */, #VALUE!,
/* |    | 7 */        "calculation"
/* |    | | */        // calculation
/* |    | 7 */        // calculation
/* |    5   */    )
/* 1        */);

// Return error message one validation style
/* 1        */RETURN_ERROR_MESSAGE_ONE_VALIDATION_STYLE = LAMBDA(
/* |2       */    numberArgument,
/* |    5   */    IF(
/* |    |6  */        NOT(ISNUMBER(numberArgument))/* numberArgument validation */, "numberArgument is invalid.",
/* |    | 7 */        "calculation"
/* |    | | */        // calculation
/* |    | 7 */        // calculation
/* |    5   */    )
/* 1        */);

// No validations style
/* 1        */NO_VALIDATIONS_STYLE = LAMBDA(
/* |      7 */    "calculation"
/* |      | */    // calculation
/* |      7 */    // calculation
/* 1        */);
```

### 6. Module structure

#### Keeping the referential transparent

Any cell references **MUST NOT** be inside function. Pass through the arguments instead.

Reason: Cell references inside function shall cause unexpected behavior and loss of testability and portability.

```excel
// Good
IS_FLUIT = LAMBDA(
    fluitName,
    fluitList, // The required information needs to be passed through the arguments.
    COUNTIF(fluitList, fluitName) > 0
);

// Bad
IS_FLUIT = LAMBDA(
    fluitName,
    COUNTIF(FluitList!$A:$A, fluitName) > 0 // A cell reference is contained.
);
```

Functions that depend on the execution environment, such as `NOW` and `TODAY`, **SHOULD NOT** be used inside function.

Reason: It isn't testable because the return value is different each time.

Exceptions: When the purpose is to extend functions that depend on the execution environment.

```excel
// Good
IS_LATE = LAMBDA(
    referenceTime, // The required information needs to be passed through the arguments.
    HOUR(referenceTime) > 9
);

// Bad
IS_LATE = LAMBDA(
    HOUR(NOW()) > 9 // A call to the `NOW` function is contained.
);
```
