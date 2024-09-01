# Excel数式スタイルガイド

## 1. 導入

### 1.1 スタイルガイドバージョン

1.0 draft

### 1.2 擁護に関する注意事項

このドキュメントでは、次の**太字かつ大文字**で修飾されたキーワードは[RFC 2119](https://datatracker.ietf.org/doc/html/rfc2119)で定義された要求レベルです。

- **しなければならない(MUST)**、**要求されている(REQUIRED)**、**することになる(SHALL)**
- **してはならない(MUST NOT)**、**することはない(SHALL NOT)**
- **する必要がある(SHOULD)**、**推奨される(RECOMMENDED)**
- **しないほうがよい(SHOULD NOT)**、**推奨されない(NOT RECOMMENDED)**
- **してもよい(MAY)**、**選択できる(OPTIONAL)**

## 2. エディタ

Advanced Formula Environment(AFE) を使用**しなければならない(MUST)**。

理由: このスタイルガイドはAFEの使用に最適化している。

備考: AFEは現在[Excel Labs](https://github.com/microsoft/Excel-Labs/)に含まれている。

## 3. 書式

### 3.1 インデント

4つの半角スペース(U+0020)を使用**しなければならない(MUST)**。

理由: AFEの標準である。

### 3.2 行の文字数

半角文字で110字以内が**推奨される(REQUIRED)**。

理由: Full-HDモニタに100%表示すると、最大化したAFEの名前及びモジュールエディタに概ね半角114文字が表示される。

例外: 改行が可読性を損なう場合。

### 3.3 カンマの配置

後置を使用**しなければならない(MUST)**。

理由: `AND`、`OR`、`IFS`、`LET`、`SWITCH`等の関数内で、式の先頭が整列される。

```excel
// 良い例
=AND(
    boolFormula1,
    boolFormula2,
    boolFormula3
)

// 悪い例 | 前置方式
=AND(
    boolFormula1 // 整列されていない。
    , boolFormula2
    , boolFormula3
)
```

### 3.4 空行

#### 3.4.1 式の内部

**推奨されない(NOT RECOMMENDED)**。

理由: Excel数式は手続き型言語ではないので、行の位置関係は意味を持つべきではない。

```excel
// 良い例
=LET(
    variable1, "値1",
    variable2, "値2",
    variable3, "値3",
    "計算処理"
)

// 悪い例
=LET(
    variable1, "値1",
    variable2, "値2",

    variable3, "値3", // 上の行は削除する。
    "計算処理"
)
```

#### 3.4.2 モジュール内の名前付きオブジェクトの間

##### Range、定数、数式

1つ以上の空行を入れ**てもよい(MAY)**。

理由: 名前付きオブジェクトの宣言は手続き的であり、宣言の位置関係に意味があることがある。

```excel
// 良い例
EXAMPLEMODULE.MODULENAME = "EXAMPLEMODULE";
EXAMPLEMODULE.MODULEVERSION = "1.0.0";

EXAMPLEMODULE.TAXRATE = 0.1;
```

##### 関数

1つ以上の空行を入れ**なければならない(MUST)**。

理由: 関数の定義はしばしば長くなる。

```excel
// 良い例
EXAMPLEMODULE.EXAMPLEFUNCTION1 = LAMBDA(
    "計算処理"
);

EXAMPLEMODULE.EXAMPLEFUNCTION2 = LAMBDA(
    "計算処理"
);
```

```excel
// 悪い例
EXAMPLEMODULE.EXAMPLEFUNCTION1 = LAMBDA(
    "計算処理"
);
EXAMPLEMODULE.EXAMPLEFUNCTION2 = LAMBDA( // ここに空行を入れる。
    "計算処理"
);
```

### 3.5 改行

改行の規則は次のとおりである:

1. 改行は演算子の後ろで**なければならず(MUST)**、続く数式にはインデントが追加され**なければならない(MUST)**。

    ```excel
    // 良い例
    =LAMBDA(
        someLongForluma1 +
            someLongForluma2 -
            someLongForluma3
    )

    // 悪い例 | 改行が演算子の前
    =LAMBDA(
        someLongForluma1
            + someLongForluma2
            - someLongForluma3
    )
    
    // 悪い例 | インデントがない
    =LAMBDA(
        someLongForluma1 +
        someLongForluma2 -
        someLongForluma3
    )
    ```

2. 左丸括弧 `(` と対応する右丸括弧`)`は同じインデントレベルで**なければならず(MUST)** 、丸括弧の間のインデントレベルは丸括弧より深く**なければならない(MUST)**。

    ```excel
    // 良い例
    =AND(
        boolFormula1,
        boolFormula2,
        boolFormula3
    )

    // 悪い例 | 左右の括弧が異なるインデントレベルにある。
    =AND(boolFormula1,
        boolFormula2,
        boolFormula3)
    
    // 悪い例 | 括弧の間のインデントレべルが括弧より深くない。
    =AND(
    boolFormula1,
    boolFormula2,
    boolFormula3
    )
    ```

### 3.6 半角スペース

1. 行末の半角スペースは存在**してはならない(MUST NOT)**。

    ```excel
    ="何らかの数式"/* Good */
    ="何らかの数式"     /* Bad */
    ```

2. 引数リストや配列区切り文字としてのカンマ`,`とセミコロン`;`は、以下の規則に従って半角スペースを持つ:
    - 次の引数が存在するならば、後ろに1つの半角スペースを持た**なければならない(MUST)**が、前に持って**はならない(MUST NOT)**。

        ```excel
        =LET(
            variableA, "値A", // 良い例
            variableB,"値B", // 悪い例
            variableC ,"値C", // 悪い例
            "計算処理"
        )
        ={1, 2, 3; 4, 5, 6; 7, 8, 9} // 良い例
        ={1,2,3;4,5,6;7,8,9} // 悪い例
        ={1 ,2 ,3 ;4 ,5 ,6 ;7 ,8 ,9} // 悪い例
        ```

    - 次の引数が省略された場合、次のカンマ`,`や右丸括弧`)`との間に半角スペースを持っ**てはならない(MUST NOT)**。

        ```excel
        =INDEX(A1:C3,, 3) // 良い例
        =INDEX(A1:C3, , 3) // 悪い例
        =INDEX(A1:C3,2,) // 良い例
        =INDEX(A1:C3,2, ) // 悪い例
        ```

3. 演算子は以下の規則に従って半角スペースを持つ:
    - 次の演算子は前後に1つの半角スペースを持た**なければならない(MUST)**。
        - 二項演算子としての`+`と`-`
        - 全ての比較演算子
        - `&`

        ```excel
        =1 + 2 // 良い例
        =1+2 // 悪い例
        =1 = 2 // 良い例
        =1=2 // 悪い例
        ="a" & "b" // 良い例
        ="a"&"b" // 悪い例
        ```

    - 次の演算子は前後に半角スペースを持っ**てはならない(MUST NOT)**。
        - `^`

        ```excel
        =1^2 // 良い例
        =1 ^ 2 // 悪い例
        ```

    - 次の演算子は前後に1つの演算子を持っ**てもよい(MAY)**。
        - `*`、`/`

        ```excel
        =1 * 2 // 良い例
        =1*2 // 良い例
        =1 / 2 // 良い例
        =1/2 // 良い例
        =1 / 2 // 良い例
        =1/2 // 良い例
        ```

4. コメントは以下の規則に従って半角スペースを持つ:
    - 行末コメント記号`//`の前には1つ、後ろには1つ以上の半角スペースを持た**なければならない(MUST)**。
    - 開ブロックコメント記号`/*`の後ろと閉ブロックコメント記号`*/`の前には1つ以上の 半角スペースを持た**なければならない(MUST)**。

    ```excel
    ="何らかの数式" // 良い例
    ="何らかの数式"//悪い例
    ="何らかの数式"/* 良い例 */
    ="何らかの数式"/*悪い例*/
    ```

5. 代入演算子としての等号`=`は前後に1つのスペースを持た**なければならない(MUST)**。

    ```excel
    // 良い例
    EXAMPLEMODULE.EXAMPLEFUNCTION = LAMBDA(
        "計算処理"
    );

    // 悪い例
    EXAMPLEMODULE.EXAMPLEFUNCTION=LAMBDA(
        "計算処理"
    );
    ```

## 4. 命名

### 4.1 全ての識別子に共通する規則

- 半角のアルファベット、数字、アンダースコア`_`のみを使用**しなければならない(MUST)**。
- 一文字目にアルファベットを使用**しなければならない(MUST)**。
- 意味のある名称が**推奨される(RECOMMENDED)**。
- 一般的でない略称は避ける**必要がある(SHOULD)**。

```excel
// 良い例
=LET(
    customerName, "名無しの権兵衛", // 意味があり、略語ではない。
    customerID, "12345", // "ID"はIDentifierの略語だが、良く知られている。
    "計算処理"
)

// 悪い例
=LET(
    a, "名無しの権兵衛", // "a"に意味はない。
    name, "名無しの権兵衛", // 意味はあるが、"name"だけではあいまいで不十分である。
    cstNm, "名無しの権兵衛", //意味はあるが、"cst"と"nm"は良く知られた略語ではない。
    "計算処理"
)
```

### 4.2 `LET`関数内の変数名

`LET`関数の変数名は次の通りである:

#### 4.2.1 局所値(Range、定数または数式)

lowerCamelCaseを使用**しなければならない(MUST)**。例えば、`customerName`。

理由: 公式の[`LET`関数リファレンス](https://support.microsoft.com/en-us/office/let-function-34842dd8-b92b-4d3f-b325-b8b8f9908999)のサンプルコードがそうであるため。

```excel
// 良い例
=LET(
    lowerCamelCase, "ローワーキャメルケース",
    "計算処理"
)

// 悪い例
=LET(
    nocapitalized, "すべて小文字",
    ALLCAPITALIZED, "すべて大文字",
    UpperCamelCase, "アッパーキャメルケース",
    snake_case, "スネークケース",
    SCREAMING_SNAKE_CASE, "スクリーミングスネークケース",
    "計算処理"
)
```

#### 4.2.2 局所関数

1. snake_caseを使用**しなければならない(MUST)**。例えば、 `calculate_price`。
2. 名前が単一の単語からなる場合、アンダースコア`_`で終端**しなければならない(MUST)**。

理由:

1. 大域関数と見分け易くするため。
2. Excelの組み込み関数名との名前の衝突を避けるため。それらの名前はアンダースコア`_`を含まない。

```excel
// 良い例
=LET(
    snake_case, LAMBDA("スネークケース"),
    "計算処理"
)

// 悪い例
=LET(
    nocapitalized, LAMBDA("すべて小文字"),
    ALLCAPITALIZED, LAMBDA("すべて大文字"),
    lowerCamelCase, LAMBDA("ローワーキャメルケース"),
    UpperCamelCase, LAMBDA("アッパーキャメルケース"),
    CALCULATE_PRICE, LAMBDA("スクリーミングスネークケース"),
    "計算処理"
)
```

### 4.3 `LAMBDA`関数内の引数名

`LET`関数内の局所変数名に準ずる**ことになる(SHALL)**。

### 4.4 単一の名前付きオブジェクト

名前付きオブジェクト(Range、定数、数式、関数)の名前はアンダースコア`_`を伴う接頭辞を持た**なければならない(MUST)**。例えば、接頭辞`EXAMPLECOMPANY`と基本となる関数名`CALCULATEPRICE`は`EXAMPLECOMPANY_CALCULATEPRICE`とする。

理由: 将来の名前の衝突を避けるため。特に全てのExcel組み込み関数はアンダースコア`_`を含まない。

名前付きオブジェクトの名前はピリオド`.`を含んで**はならない(MUST NOT)**。

理由: 同じ名前空間に含まれる名前付きオブジェクトは常に同時にインポートされるべきであるが、単一オブジェクトでは出来ないため。

### 4.5 ワークブックモジュール内の名前付きオブジェクト

名前付きオブジェクト(Range、定数、数式、関数)の名前はアンダースコア`_`を伴う接頭辞を持た**なければならない(MUST)**。

理由: 単一の名前付きオブジェクトと同じ理由である。

名前付きオブジェクトの名前はピリオド`.`を含ま**ないほうがよい(SHOULD NOT)**。

理由: 将来の移植性を維持するため。ピリオド `.`はワークブックモジュールでは許可されているが、その他のモジュールでは許可されていない。

例外: 名前空間区切り文字としてのピリオド`.`を1つだけ使用する場合。

### 4.6 その他のモジュール内の名前付きオブジェクト

モジュール名とモジュール内の名前付きオブジェクト名(Range、定数、数式、関数) はアンダースコア`_`を含んで**もよい(MAY)**。

理由: Excel組み込み関数は名前空間を使用するが、僅かであるため名前の衝突リスクが低い。

### 5. 関数の構造

#### 5.1 全体構造

関数は以下のように構造化する:

1. 関数定義全体。
2. 引数(オプション)。
3. 引数検証を伴う計算処理部分。ここは`LET`関数で囲む**必要がある(SHOULD)**。  
    例外: 引数検証が無いか、引数が0または1つの場合。
4. 引数検証の結果を代入するための変数定義と引数検証。
5. 引数検証のガード節を伴う計算処理部分。ここは`IF`関数で囲む**必要がある(SHOULD)**。  
    例外: 引数検証が無い場合。
6. 引数検証のガード節。引数検証の結果を検証し、無効であれば`#VALUE!`エラーかエラーメッセージを返さ**なければならない(MUST)**。
7. 主たる計算処理。

```excel
// エラーを発生させる基本スタイル
/* 1        */RAISE_ERROR_BASIC_STYLE = LAMBDA(
/* |2       */    numberArgument,
/* ||       */    optionaltextArgument,
/* |2       */    enumerateTextArgument,
/* | 3      */    LET(
/* | |4     */        argumentIsInvalid, OR(
/* | ||     */            OR(ISOMITTED(numberArgument), NOT(ISNUMBER(numberArgument))), // numberArgumentの検証
/* | ||     */            AND(NOT(ISOMITTED(optionaltextArgument)), NOT(ISTEXT(optionaltextArgument))), // optionaltextArgumentの検証
/* | ||     */            ISERROR(LET(list, {"りんご", "ばなな", "みかん"},FILTER(list, list=enumerateTextArgument))) // enumerateTextArgumentの検証
/* | |4     */        ),
/* | |  5   */        IF(
/* | |  |6  */            argumentIsInvalid, #VALUE!,
/* | |  | 7 */            "計算処理"
/* | |  | | */            // 計算処理
/* | |  | 7 */            // 計算処理
/* | |  5   */        )
/* | 3      */    )
/* 1        */);

// エラーメッセージを返す基本スタイル
/* 1        */RETURN_ERROR_MESSAGE_BASIC_STYLE = LAMBDA(
/* |2       */    numberArgument,
/* ||       */    optionaltextArgument,
/* |2       */    enumerateTextArgument,
/* | 3      */    LET(
/* | |4     */        errorMessage, IFS(
/* | ||     */            OR(ISOMITTED(numberArgument), NOT(ISNUMBER(numberArgument))), "numberArgument is invalid.", // numberArgumentの検証
/* | ||     */            AND(NOT(ISOMITTED(optionaltextArgument)), NOT(ISTEXT(optionaltextArgument))), "optionaltextArgument is invalid.", // optionaltextArgumentの検証
/* | ||     */            ISERROR(LET(list, {"りんご", "ばなな", "みかん"},FILTER(list, list=enumerateTextArgument))), "enumerateTextArgument is invalid.", // enumerateTextArgumentの検証
/* | ||     */            TRUE, ""
/* | |4     */        ),
/* | |  5   */        IF(
/* | |  |6  */            errorMessage <> "", errorMessage,
/* | |  | 7 */            "計算処理"
/* | |  | | */            // 計算処理
/* | |  | 7 */            // 計算処理
/* | |  5   */        )
/* | 3      */    )
/* 1        */);

// エラーを発生させる1つの検証スタイル
/* 1        */RAISE_ERROR_ONE_VALIDATION_STYLE = LAMBDA(
/* |2       */    numberArgument,
/* |    5   */    IF(
/* |    |6  */        NOT(ISNUMBER(numberArgument))/* numberArgumentの検証 */, #VALUE!,
/* |    | 7 */        "計算処理"
/* |    | | */        // 計算処理
/* |    | 7 */        // 計算処理
/* |    5   */    )
/* 1        */);

// エラーメッセージを返す1つの検証スタイル
/* 1        */RETURN_ERROR_MESSAGE_ONE_VALIDATION_STYLE = LAMBDA(
/* |2       */    numberArgument,
/* |    5   */    IF(
/* |    |6  */        NOT(ISNUMBER(numberArgument))/* numberArgumentの検証 */, "numberArgument is invalid.",
/* |    | 7 */        "計算処理"
/* |    | | */        // 計算処理
/* |    | 7 */        // 計算処理
/* |    5   */    )
/* 1        */);

// 検証無しスタイル
/* 1        */NO_VALIDATIONS_STYLE = LAMBDA(
/* |      7 */    "計算処理"
/* |      | */    // 計算処理
/* |      7 */    // 計算処理
/* 1        */);
```

### 6. モジュールの構造

#### 参照透明性の維持

関数内にあらゆるセル参照を含んで**はならない(MUST NOT)**。代替として引数渡しを用いる。

理由: 関数内のセル参照は意図せぬふるまいを引き起こし、テスト性と移植性を損なう。

```excel
// 良い例
IS_FLUIT = LAMBDA(
    fluitName,
    fluitList, // 必要な情報は引数渡しされる必要がある。
    COUNTIF(fluitList, fluitName) > 0
);

// 悪い例
IS_FLUIT = LAMBDA(
    fluitName,
    COUNTIF(FluitList!$A:$A, fluitName) > 0 // セル参照が含まれている。
);
```

`NOW`や`TODAY`のような実行環境に依存する関数は関数内で使用**しないほうがよい(SHOULD NOT)**。

理由: 戻り値が毎回異なるためテストしづらい。

例外: 実行環境に依存する関数の拡張自体が目的である場合。

```excel
// 良い例
IS_LATE = LAMBDA(
    referenceTime, // 必要な情報は引数渡しされる必要がある。
    HOUR(referenceTime) > 9
);

// 悪い例
IS_LATE = LAMBDA(
    HOUR(NOW()) > 9 // `NOW`関数の呼び出しが含まれている。
);
```
