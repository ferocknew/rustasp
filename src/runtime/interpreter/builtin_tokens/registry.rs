/// Token 注册表
/// 维护函数名到 Token ID 的映射
pub struct TokenRegistry {
    /// 函数名到 Token 的映射（小写存储，不区分大小写）
    map: HashMap<String, BuiltinToken>,
}

impl TokenRegistry {
    /// 创建新的 Token 注册表并初始化所有内置函数
    pub fn new() -> Self {
        let mut registry = Self {
            map: HashMap::new(),
        };
        registry.init_all_tokens();
        registry
    }

    /// 初始化所有内置函数 Token
    fn init_all_tokens(&mut self) {
        // 数学函数
        self.register("abs", BuiltinToken::Abs);
        self.register("sqr", BuiltinToken::Sqr);
        self.register("sin", BuiltinToken::Sin);
        self.register("cos", BuiltinToken::Cos);
        self.register("tan", BuiltinToken::Tan);
        self.register("atn", BuiltinToken::Atn);
        self.register("log", BuiltinToken::Log);
        self.register("exp", BuiltinToken::Exp);
        self.register("int", BuiltinToken::Int);
        self.register("fix", BuiltinToken::Fix);
        self.register("round", BuiltinToken::Round);
        self.register("rnd", BuiltinToken::Rnd);
        self.register("sgn", BuiltinToken::Sgn);

        // 类型转换函数
        self.register("cstr", BuiltinToken::CStr);
        self.register("cint", BuiltinToken::CInt);
        self.register("clng", BuiltinToken::CLng);
        self.register("csng", BuiltinToken::CSng);
        self.register("cdbl", BuiltinToken::CDbl);
        self.register("cbool", BuiltinToken::CBool);
        self.register("cdate", BuiltinToken::CDate);
        self.register("cbyte", BuiltinToken::CByte);

        // 字符串函数
        self.register("len", BuiltinToken::Len);
        self.register("trim", BuiltinToken::Trim);
        self.register("ltrim", BuiltinToken::LTrim);
        self.register("rtrim", BuiltinToken::RTrim);
        self.register("left", BuiltinToken::Left);
        self.register("right", BuiltinToken::Right);
        self.register("mid", BuiltinToken::Mid);
        self.register("ucase", BuiltinToken::UCase);
        self.register("lcase", BuiltinToken::LCase);
        self.register("instr", BuiltinToken::InStr);
        self.register("instrrev", BuiltinToken::InStrRev);
        self.register("strcomp", BuiltinToken::StrComp);
        self.register("replace", BuiltinToken::Replace);
        self.register("split", BuiltinToken::Split);
        self.register("join", BuiltinToken::Join);
        self.register("strreverse", BuiltinToken::StrReverse);
        self.register("space", BuiltinToken::Space);
        self.register("string", BuiltinToken::String_);
        self.register("asc", BuiltinToken::Asc);
        self.register("chr", BuiltinToken::Chr);
        self.register("ascw", BuiltinToken::AscW);
        self.register("chrw", BuiltinToken::ChrW);

        // 日期时间函数
        self.register("now", BuiltinToken::Now);
        self.register("date", BuiltinToken::Date);
        self.register("time", BuiltinToken::Time);
        self.register("year", BuiltinToken::Year);
        self.register("month", BuiltinToken::Month);
        self.register("day", BuiltinToken::Day);
        self.register("hour", BuiltinToken::Hour);
        self.register("minute", BuiltinToken::Minute);
        self.register("second", BuiltinToken::Second);
        self.register("weekday", BuiltinToken::WeekDay);
        self.register("dateadd", BuiltinToken::DateAdd);
        self.register("datediff", BuiltinToken::DateDiff);
        self.register("datepart", BuiltinToken::DatePart);
        self.register("dateserial", BuiltinToken::DateSerial);
        self.register("datevalue", BuiltinToken::DateValue);
        self.register("timeserial", BuiltinToken::TimeSerial);
        self.register("timevalue", BuiltinToken::TimeValue);
        self.register("formatdatetime", BuiltinToken::FormatDateTime);
        self.register("monthname", BuiltinToken::MonthName);
        self.register("weekdayname", BuiltinToken::WeekDayName);

        // 数组函数
        self.register("array", BuiltinToken::Array);
        self.register("ubound", BuiltinToken::UBound);
        self.register("lbound", BuiltinToken::LBound);
        self.register("filter", BuiltinToken::Filter);
        self.register("isarray", BuiltinToken::IsArray);

        // 检验函数
        self.register("isnumeric", BuiltinToken::IsNumeric);
        self.register("isdate", BuiltinToken::IsDate);
        self.register("isempty", BuiltinToken::IsEmpty);
        self.register("isnull", BuiltinToken::IsNull);
        self.register("isobject", BuiltinToken::IsObject);
        self.register("isnothing", BuiltinToken::IsNothing);
        self.register("typename", BuiltinToken::TypeName);
        self.register("vartype", BuiltinToken::VarType);

        // 交互函数
        self.register("msgbox", BuiltinToken::MsgBox);
        self.register("inputbox", BuiltinToken::InputBox);

        // 其他函数
        self.register("createobject", BuiltinToken::CreateObject);
        self.register("getobject", BuiltinToken::GetObject);
        self.register("eval", BuiltinToken::Eval);
        self.register("execute", BuiltinToken::Execute);
        self.register("rgb", BuiltinToken::RGB);
    }

    /// 注册单个 Token
    fn register(&mut self, name: &str, token: BuiltinToken) {
        self.map.insert(name.to_lowercase(), token);
    }

    /// 查找函数名对应的 Token
    pub fn lookup(&self, name: &str) -> Option<BuiltinToken> {
        self.map.get(&name.to_lowercase()).copied()
    }

    /// 检查是否为内置函数
    pub fn is_builtin(&self, name: &str) -> bool {
        self.map.contains_key(&name.to_lowercase())
    }
}

impl Default for TokenRegistry {
    fn default() -> Self {
        Self::new()
    }
}
