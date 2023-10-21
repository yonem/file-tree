## 動作環境

- Java21で動作

2023/9/29現在  
https://www.oracle.com/java/technologies/downloads/#java21

- JUnit5でテスト

```xml

<dependency>
    <groupId>org.junit.jupiter</groupId>
    <artifactId>junit-jupiter-api</artifactId>
    <version>5.9.3</version>
    <scope>test</scope>
</dependency>
```

## 操作方法

1. 起動時にディレクトリ選択のダイアログを表示
2. ツリーの起点となるディレクトリを選択する
3. 選択したディレクトリの直下に収集結果が出力される

![出力イメージ](img/output_image.png)
