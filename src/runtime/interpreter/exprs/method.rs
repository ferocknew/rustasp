//! 方法调用表达式求值

use crate::ast::Expr;
use crate::runtime::{RuntimeError, Value, VbsInstance, ControlFlow};

use super::super::Interpreter;

impl Interpreter {
    /// 执行方法调用
    pub fn eval_method(
        &mut self,
        object: &Expr,
        method: &str,
        args: &[Expr],
    ) -> Result<Value, RuntimeError> {
        // 计算参数值
        let arg_values: Result<Vec<Value>, _> =
            args.iter().map(|e| self.eval_expr(e)).collect();
        let arg_values = arg_values?;

        let method_lower = method.to_lowercase();

        // 特殊处理：Err 对象方法调用
        if let Expr::Variable(name) = object {
            if name.to_lowercase() == "err" {
                return self.eval_err_method(&method_lower);
            }
        }

        // 先计算对象表达式
        let obj_val = self.eval_expr(object)?;

        // 处理类实例的方法调用
        if let Value::Object(ref obj) = obj_val {
            // 尝试向下转型为 VbsInstance
            let locked = obj.lock()
                .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?;

            if let Some(instance) = locked.as_any().downcast_ref::<VbsInstance>() {
                // 获取类名
                let class_name = instance.class_name.clone();
                let normalized_class = crate::utils::normalize_identifier(&class_name);

                // 释放锁后再获取类定义（避免借用冲突）
                drop(locked);

                // 从上下文获取类定义
                let vbs_class = self.context.classes.get(&normalized_class).cloned();
                if let Some(vbs_class) = vbs_class {
                    // 查找方法（大小写不敏感）
                    let method_decl = vbs_class.methods.iter()
                        .find(|(name, _)| crate::utils::normalize_identifier(name) == method_lower)
                        .map(|(_, decl)| decl.clone());

                    if let Some(method_decl) = method_decl {
                        // 获取实例字段作为参数绑定
                        let instance_fields = {
                            let locked = obj.lock()
                                .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?;
                            if let Some(inst) = locked.as_any().downcast_ref::<VbsInstance>() {
                                inst.fields.clone()
                            } else {
                                return Err(RuntimeError::Generic("Invalid instance".to_string()));
                            }
                        };

                        // 执行方法（使用已计算的 arg_values）
                        return self.execute_class_method_with_values(&method_decl, args, arg_values, instance_fields, obj.clone());
                    }
                }

                return Err(RuntimeError::MethodNotFound(format!("{}.{}", class_name, method)));
            }

            // 释放锁
            drop(locked);

            // 内置对象方法调用（使用已计算的参数值）
            let result = obj.lock()
                .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?
                .call_method(&method_lower, arg_values);

            return result;
        }

        Ok(Value::Empty)
    }

    /// 执行类方法（供 property.rs 调用）
    pub(super) fn execute_class_method(
        &mut self,
        method_decl: &crate::ast::MethodDecl,
        args: &[Expr],
        instance_fields: std::collections::HashMap<String, Value>,
        instance: crate::runtime::ObjectRef,
    ) -> Result<Value, RuntimeError> {
        // 计算参数值
        let arg_values: Result<Vec<Value>, _> =
            args.iter().map(|e| self.eval_expr(e)).collect();
        let arg_values = arg_values?;

        self.execute_class_method_with_values(method_decl, args, arg_values, instance_fields, instance)
    }

    /// 执行类方法（使用已计算的参数值）
    fn execute_class_method_with_values(
        &mut self,
        method_decl: &crate::ast::MethodDecl,
        _args: &[Expr],
        arg_values: Vec<Value>,
        instance_fields: std::collections::HashMap<String, Value>,
        instance: crate::runtime::ObjectRef,
    ) -> Result<Value, RuntimeError> {
        // 推入新作用域
        self.context.push_scope();

        // 绑定 Me 关键字（指向当前实例）
        self.context.define_var("Me".to_string(), Value::Object(instance.clone()));

        // 绑定参数
        for (i, param) in method_decl.params.iter().enumerate() {
            let value = if i < arg_values.len() {
                arg_values[i].clone()
            } else {
                Value::Empty
            };
            self.context.define_var(param.name.clone(), value);
        }

        // 初始化方法名变量为 Empty（用于返回值）
        let method_name_lower = crate::utils::normalize_identifier(&method_decl.name);
        self.context.define_var(method_decl.name.clone(), Value::Empty);

        // 将实例字段复制到当前作用域（支持直接访问字段）
        for (name, value) in &instance_fields {
            self.context.define_var(name.clone(), value.clone());
        }

        // 执行方法体
        for stmt in &method_decl.body {
            match self.eval_stmt(stmt) {
                Ok(_) => {}
                Err(RuntimeError::ControlFlow(ControlFlow::ExitFunction)) |
                Err(RuntimeError::ControlFlow(ControlFlow::ExitSub)) => {
                    // Exit Function/Sub - 正常退出
                    break;
                }
                Err(e) => {
                    self.context.pop_scope();
                    return Err(e);
                }
            }
        }

        // 获取返回值（Function 方法名变量的值）
        let result = self.context.get_var(&method_decl.name)
            .or_else(|| self.context.get_var(&method_name_lower))
            .cloned()
            .unwrap_or(Value::Empty);

        // 同步实例字段（方法可能修改了字段）
        {
            let mut locked = instance.lock()
                .map_err(|_| RuntimeError::Generic("Failed to lock object".to_string()))?;
            if let Some(inst) = locked.as_any_mut().downcast_mut::<VbsInstance>() {
                for name in instance_fields.keys() {
                    if let Some(value) = self.context.get_var(name).cloned() {
                        let _ = inst.set_field(name.clone(), value);
                    }
                }
            }
        }

        self.context.pop_scope();

        Ok(result)
    }

    /// 处理 Err 对象的方法调用
    fn eval_err_method(&mut self, method: &str) -> Result<Value, RuntimeError> {
        match method {
            "clear" => {
                self.context.err.clear();
                Ok(Value::Empty)
            }
            _ => Err(RuntimeError::MethodNotFound(format!("Err.{}", method))),
        }
    }
}
