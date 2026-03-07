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

        // 解析类成员直到 EndClass
        while !self.check_keyword(Keyword::EndClass) && !self.check(&Token::Eof) {
            if let Some(member) = self.parse_class_member()? {
                members.push(member);
            }
        }

        self.expect_keyword(Keyword::EndClass)?;

        Ok(Some(Stmt::Class { name, members }))
    }

    /// 解析类成员
    fn parse_class_member(&mut self) -> Result<Option<ClassMember>, ParseError> {
        // 解析可见性修饰符（Public/Private）
        let visibility = self.parse_visibility();

        match self.peek() {
            // Function 或 Sub - 方法成员
            Token::Keyword(Keyword::Function) => {
                let method = self.parse_method(visibility)?;
                Ok(Some(ClassMember::Method(method)))
            }
            Token::Keyword(Keyword::Sub) => {
                let method = self.parse_method(visibility)?;
                Ok(Some(ClassMember::Method(method)))
            }

            // Property Get/Let/Set - 属性成员
            Token::Keyword(Keyword::Property) => {
                let property = self.parse_property(visibility)?;
                Ok(Some(ClassMember::Property(property)))
            }

            // Dim 或 标识符 - 字段成员
            Token::Keyword(Keyword::Dim) => {
                let field = self.parse_field(visibility)?;
                Ok(Some(ClassMember::Field(field)))
            }

            // 标识符 - 简写字段声明（没有 Dim）
            Token::Ident(_) => {
                let field = self.parse_field(visibility)?;
                Ok(Some(ClassMember::Field(field)))
            }

            // 跳过空行
            Token::Newline => {
                self.advance();
                Ok(None)
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
    /// ```
    fn parse_field(&mut self, visibility: Visibility) -> Result<FieldDecl, ParseError> {
        // 跳过 Dim 关键字（如果有）
        if self.check_keyword(Keyword::Dim) {
            self.advance();
        }

        let name = self.expect_ident()?;

        // VBScript 类字段不支持初始化和数组维度
        // 所以这里只需要读取名称

        // 消耗换行符
        self.skip_newlines();

        Ok(FieldDecl { name, visibility })
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
                    // 检查是否是 End Function/End Sub
                    self.advance();
                    if is_function && self.check_keyword(Keyword::Function) {
                        self.advance();
                        break;
                    } else if is_sub && self.check_keyword(Keyword::Sub) {
                        self.advance();
                        break;
                    } else {
                        // 不是我们的 End，继续解析
                        continue;
                    }
                }
                Token::Keyword(Keyword::EndClass) | Token::Eof => {
                    // 未闭合的方法定义
                    break;
                }
                Token::Newline => {
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
                    self.advance();
                    if self.check_keyword(Keyword::Property) {
                        self.advance();
                        break;
                    } else {
                        // 不是我们的 End Property，继续
                        continue;
                    }
                }
                Token::Keyword(Keyword::EndClass) | Token::Eof => {
                    break;
                }
                Token::Newline => {
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
