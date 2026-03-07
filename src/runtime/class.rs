//! VBScript 类运行时支持
//!
//! 实现 VbsClass（类定义）和 VbsInstance（类实例）

use std::collections::HashMap;

use crate::ast::{ClassMember, FieldDecl, MethodDecl, PropertyDecl, PropertyType, Visibility};
use crate::runtime::{BuiltinObject, RuntimeError, Value};

/// VBScript 类定义
#[derive(Debug, Clone)]
pub struct VbsClass {
    /// 类名
    pub name: String,
    /// 字段声明（字段名 -> 声明）
    pub fields: HashMap<String, FieldDecl>,
    /// 常量声明（字段名 -> (值, 可见性)）
    pub constants: HashMap<String, (Value, Visibility)>,
    /// 方法声明（方法名 -> 声明）
    pub methods: HashMap<String, MethodDecl>,
    /// 属性声明（属性名 -> 声明）
    pub properties: HashMap<String, PropertyDecl>,
}

impl VbsClass {
    /// 从 AST 创建类定义
    pub fn from_ast(name: String, members: Vec<ClassMember>) -> Self {
        let mut fields = HashMap::new();
        let mut constants = HashMap::new();
        let mut methods = HashMap::new();
        let mut properties = HashMap::new();

        for member in members {
            match member {
                ClassMember::Field(field) => {
                    fields.insert(field.name.clone(), field);
                }
                ClassMember::Const { name, value: _, visibility } => {
                    // 常量值需要在解释阶段求值，这里先存储为 Empty
                    // 解释器会在注册类时求值这些常量
                    constants.insert(name, (Value::Empty, visibility));
                }
                ClassMember::Method(method) => {
                    methods.insert(method.name.clone(), method);
                }
                ClassMember::Property(property) => {
                    // 属性名包含类型后缀，如 Name_Get, Name_Let, Name_Set
                    let key = format!("{}_{}", property.name, Self::prop_type_suffix(&property.prop_type));
                    properties.insert(key, property);
                }
            }
        }

        VbsClass {
            name,
            fields,
            constants,
            methods,
            properties,
        }
    }

    /// 获取属性类型的后缀
    fn prop_type_suffix(prop_type: &PropertyType) -> &'static str {
        match prop_type {
            PropertyType::Get => "Get",
            PropertyType::Let => "Let",
            PropertyType::Set => "Set",
        }
    }

    /// 查找字段
    pub fn get_field(&self, name: &str) -> Option<&FieldDecl> {
        self.fields.get(name)
    }

    /// 查找方法
    pub fn get_method(&self, name: &str) -> Option<&MethodDecl> {
        self.methods.get(name)
    }

    /// 查找属性
    pub fn get_property(&self, name: &str, prop_type: &PropertyType) -> Option<&PropertyDecl> {
        let key = format!("{}_{}", name, Self::prop_type_suffix(prop_type));
        self.properties.get(&key)
    }

    /// 创建新实例
    pub fn new_instance(&self) -> VbsInstance {
        let mut instance_fields = HashMap::new();

        // 初始化所有字段为 Empty
        for (name, _field) in &self.fields {
            instance_fields.insert(name.clone(), Value::Empty);
        }

        // 初始化常量
        for (name, (value, _visibility)) in &self.constants {
            instance_fields.insert(name.clone(), value.clone());
        }

        VbsInstance {
            class_name: self.name.clone(),
            fields: instance_fields,
        }
    }
}

/// VBScript 类实例
#[derive(Debug)]
pub struct VbsInstance {
    /// 类名（用于调试和类型检查）
    pub class_name: String,
    /// 字段值（字段名 -> 值）
    pub fields: HashMap<String, Value>,
}

impl VbsInstance {
    /// 获取字段值
    pub fn get_field(&self, name: &str) -> Option<&Value> {
        self.fields.get(name)
    }

    /// 设置字段值
    pub fn set_field(&mut self, name: String, value: Value) -> Result<(), RuntimeError> {
        if self.fields.contains_key(&name) {
            self.fields.insert(name, value);
            Ok(())
        } else {
            Err(RuntimeError::Generic(format!(
                "Field '{}' not found in class '{}'",
                name, self.class_name
            )))
        }
    }

    /// 转换为 Value
    pub fn to_value(self) -> Value {
        Value::Object(Box::new(self))
    }
}

/// 为 BuiltinObject trait 实现
impl BuiltinObject for VbsInstance {
    fn clone_box(&self) -> Box<dyn BuiltinObject> {
        // VBScript 对象是引用类型，克隆返回同一个引用
        Box::new(VbsInstance {
            class_name: self.class_name.clone(),
            fields: self.fields.clone(),
        })
    }

    fn get_property(&self, name: &str) -> Result<Value, RuntimeError> {
        match self.get_field(name) {
            Some(value) => Ok(value.clone()),
            None => Err(RuntimeError::Generic(format!(
                "Property or field '{}' not found",
                name
            ))),
        }
    }

    fn set_property(&mut self, name: &str, value: Value) -> Result<(), RuntimeError> {
        self.set_field(name.to_string(), value)
    }

    fn call_method(&mut self, _name: &str, _args: Vec<Value>) -> Result<Value, RuntimeError> {
        // 方法调用在解释器层面处理，这里只是占位
        Err(RuntimeError::Generic(
            "Method call should be handled by interpreter".to_string(),
        ))
    }

    fn index(&self, key: &Value) -> Result<Value, RuntimeError> {
        // 默认实现：尝试作为属性访问
        let key_str = key.to_string();
        self.get_property(&key_str)
    }

    fn as_any(&self) -> &dyn std::any::Any {
        self
    }

    fn as_any_mut(&mut self) -> &mut dyn std::any::Any {
        self
    }
}
