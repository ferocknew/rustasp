//! Class 类解析器

use crate::ast::{ClassMember, FieldDecl, MethodDecl, PropertyDecl, PropertyType, Stmt, Visibility};
use crate::parser::keyword::Keyword;
use crate::parser::lexer::Token;
use crate::parser::ParseError;
use crate::parser::Parser;

impl Parser {
    /// 解析 Class 定义
    ///
    /// 语法:
    /// ```vbscript
    /// Class Person
    ///     Public Name
    ///     Private age
    ///     Public Function GetAge()
    ///         GetAge = age
    ///     End Function
    /// End Class
    /// ```
    pub fn parse_class(&mut self) -> Result<Option<Stmt>, ParseError> {
        self.expect_keyword(Keyword::Class)?;

        let name = self.expect_ident()?;

        let mut members = Vec::new();

        // 解析类成员直到 End Class
        loop {
            self.skip_newlines();
            if self.check(&Token::Eof) {
                return Err(ParseError::UnexpectedEof("expected End Class".to_string()));
            }
            // 检查是否是 "End Class" (两个关键字)
            if self.check_keyword(Keyword::End) {
                // 向前看下一个 token 是否是 Class
                if self.peek_next_is_keyword(Keyword::Class) {
                    break; // 找到 End Class，退出循环
                }
            }
            let new_members = self.parse_class_member()?;
            members.extend(new_members);
        }

        // 消耗 "End Class"
        self.expect_keyword(Keyword::End)?;
        self.expect_keyword(Keyword::Class)?;

        Ok(Some(Stmt::Class { name, members }))
    }

    /// 检查下一个 token 是否是指定关键字
    fn peek_next_is_keyword(&self, keyword: Keyword) -> bool {
        let next_token = self.peek_ahead(1);
        matches!(next_token, Token::Keyword(k) if *k == keyword)
    }

    /// 解析类成员
    fn parse_class_member(&mut self) -> Result<Vec<ClassMember>, ParseError> {
        // 先检查是否到达 End Class（严格检查 End + Class）
        if self.check_keyword(Keyword::End) && self.peek_next_is_keyword(Keyword::Class) {
            return Ok(vec![]);  // 返回空，让 parse_class 处理 End Class
        }

        // 解析可见性修饰符（Public/Private）
        let visibility = self.parse_visibility();

        match self.peek() {
            // Function 或 Sub - 方法成员
            Token::Keyword(Keyword::Function) => {
                let method = self.parse_method(visibility)?;
                Ok(vec![ClassMember::Method(method)])
            }
            Token::Keyword(Keyword::Sub) => {
                let method = self.parse_method(visibility)?;
                Ok(vec![ClassMember::Method(method)])
            }

            // Property Get/Let/Set - 属性成员
            Token::Keyword(Keyword::Property) => {
                let property = self.parse_property(visibility)?;
                Ok(vec![ClassMember::Property(property)])
            }

            // Dim 或 标识符 - 字段成员
            Token::Keyword(Keyword::Dim) => {
                let fields = self.parse_field(visibility)?;
                Ok(fields.into_iter().map(ClassMember::Field).collect())
            }

            // 标识符 - 简写字段声明（没有 Dim）
            Token::Ident(_) => {
                let fields = self.parse_field(visibility)?;
                Ok(fields.into_iter().map(ClassMember::Field).collect())
            }

            // 跳过空行
            Token::Newline => {
                self.advance();
                Ok(vec![])
            }

            _ => {
                let token = self.peek();
                Err(ParseError::UnexpectedToken {
                    expected: "class member (Function, Sub, Property, Dim, or identifier)".to_string(),
                    found: format!("{:?}", token),
                })
            }
        }
    }

    /// 解析可见性修饰符
    fn parse_visibility(&mut self) -> Visibility {
        if self.check_keyword(Keyword::Public) {
            self.advance();
            Visibility::Public
        } else if self.check_keyword(Keyword::Private) {
            self.advance();
            Visibility::Private
        } else {
            // 默认为 Public（VBScript 中类成员默认是 Public）
            Visibility::Public
        }
    }

    /// 解析字段声明
    ///
    /// 语法:
    /// ```vbscript
    /// Public Name
    /// Private age
    /// Dim count
    /// Public Lang, [Error], Str, Var  ' 支持逗号分隔的多字段声明
    /// ```
    fn parse_field(&mut self, visibility: Visibility) -> Result<Vec<FieldDecl>, ParseError> {
        // 跳过 Dim 关键字（如果有）
        if self.check_keyword(Keyword::Dim) {
            self.advance();
        }

        let mut fields = Vec::new();

        // 解析第一个字段名
        let name = self.expect_ident()?;
        fields.push(FieldDecl {
            name,
            visibility: visibility.clone(),
        });

        // 检查是否有逗号分隔的更多字段
        while self.check(&Token::Comma) {
            self.advance(); // 消耗逗号
            let name = self.expect_ident()?;
            fields.push(FieldDecl {
                name,
                visibility: visibility.clone(),
            });
        }

        // 消耗换行符
        self.skip_newlines();

        Ok(fields)
    }

    /// 解析方法声明（Function 或 Sub）
    ///
    /// 语法:
    /// ```vbscript
    /// Public Function GetName()
    ///     GetName = Name
    /// End Function
    ///
    /// Public Sub SetName(value)
    ///     Name = value
    /// End Sub
    /// ```
    fn parse_method(&mut self, visibility: Visibility) -> Result<MethodDecl, ParseError> {
        let is_function = self.check_keyword(Keyword::Function);
        let is_sub = self.check_keyword(Keyword::Sub);

        if !is_function && !is_sub {
            return Err(ParseError::UnexpectedToken {
                expected: "Function or Sub".to_string(),
                found: format!("{:?}", self.peek()),
            });
        }

        self.advance(); // 消耗 Function 或 Sub

        let name = self.expect_ident()?;

        let params = self.parse_params()?;

        let mut body = Vec::new();

        // 解析方法体
        loop {
            match self.peek() {
                Token::Keyword(Keyword::End) => {
                    // 检查是否是 End Function/End Sub/End Class
                    // 注意：不要先 advance()，要先 peek 下一个 token
                    if is_function && self.peek_next_is_keyword(Keyword::Function) {
                        self.advance(); // 消耗 End
                        self.advance(); // 消耗 Function
                        break;
                    } else if is_sub && self.peek_next_is_keyword(Keyword::Sub) {
                        self.advance(); // 消耗 End
                        self.advance(); // 消耗 Sub
                        break;
                    } else if self.peek_next_is_keyword(Keyword::Class) {
                        // 遇到 End Class，方法未闭合，不要消耗 End
                        break;
                    } else {
                        // 不是我们的 End，让 parse_stmt 处理
                        if let Some(stmt) = self.parse_stmt()? {
                            body.push(stmt);
                        }
                    }
                }
                Token::Eof => {
                    // 未闭合的方法定义
                    break;
                }
                Token::Newline => {
                    self.advance();
                }
                Token::Colon => {
                    // 冒号语句分隔符
                    self.advance();
                }
                _ => {
                    if let Some(stmt) = self.parse_stmt()? {
                        body.push(stmt);
                    }
                }
            }
        }

        Ok(MethodDecl {
            name,
            params,
            body,
            visibility,
        })
    }

    /// 解析属性声明（Property Get/Let/Set）
    ///
    /// 语法:
    /// ```vbscript
    /// Property Get Name
    ///     Name = mName
    /// End Property
    ///
    /// Property Let Name(value)
    ///     mName = value
    /// End Property
    /// ```
    fn parse_property(&mut self, visibility: Visibility) -> Result<PropertyDecl, ParseError> {
        self.expect_keyword(Keyword::Property)?;

        let prop_type = if self.check_keyword(Keyword::Get) {
            self.advance();
            PropertyType::Get
        } else if self.check_keyword(Keyword::Let) {
            self.advance();
            PropertyType::Let
        } else if self.check_keyword(Keyword::Set) {
            self.advance();
            PropertyType::Set
        } else {
            return Err(ParseError::UnexpectedToken {
                expected: "Get, Let, or Set".to_string(),
                found: format!("{:?}", self.peek()),
            });
        };

        let name = self.expect_ident()?;

        let params = self.parse_params()?;

        let mut body = Vec::new();

        // 解析属性体
        loop {
            match self.peek() {
                Token::Keyword(Keyword::End) => {
                    // 注意：不要先 advance()，要先 peek 下一个 token
                    if self.peek_next_is_keyword(Keyword::Property) {
                        self.advance(); // 消耗 End
                        self.advance(); // 消耗 Property
                        break;
                    } else if self.peek_next_is_keyword(Keyword::Class) {
                        // 遇到 End Class，属性未闭合，不要消耗 End
                        break;
                    } else {
                        // 不是我们的 End Property，让 parse_stmt 处理
                        if let Some(stmt) = self.parse_stmt()? {
                            body.push(stmt);
                        }
                    }
                }
                Token::Eof => {
                    break;
                }
                Token::Newline => {
                    self.advance();
                }
                Token::Colon => {
                    // 冒号语句分隔符
                    self.advance();
                }
                _ => {
                    if let Some(stmt) = self.parse_stmt()? {
                        body.push(stmt);
                    }
                }
            }
        }

        Ok(PropertyDecl {
            name,
            params,
            body,
            visibility,
            prop_type,
        })
    }
}
