//! FileSystemObject 对象 - 文件系统操作

use crate::runtime::{BuiltinObject, RuntimeError, Value, ValueConversion};
use std::fs;
use std::path::Path;

/// FileSystemObject 对象
#[derive(Debug, Clone)]
pub struct FileSystemObject;

impl FileSystemObject {
    /// 创建新 FileSystemObject
    pub fn new() -> Self {
        FileSystemObject
    }

    /// 检查文件是否存在
    pub fn file_exists(&self, path: &str) -> bool {
        Path::new(path).exists() && Path::new(path).is_file()
    }

    /// 检查文件夹是否存在
    pub fn folder_exists(&self, path: &str) -> bool {
        Path::new(path).exists() && Path::new(path).is_dir()
    }

    /// 检查驱动器是否存在
    pub fn drive_exists(&self, drive: &str) -> bool {
        // 简化实现：假设驱动器存在（仅用于演示）
        // 实际实现需要检查系统驱动器
        !drive.is_empty()
    }

    /// 获取驱动器名称
    pub fn get_drive_name(&self, path: &str) -> String {
        // 简化实现：从路径中提取驱动器部分
        if let Some(idx) = path.find(':') {
            path[..idx + 1].to_uppercase()
        } else if path.starts_with('/') {
            "/".to_string()
        } else {
            "".to_string()
        }
    }

    /// 获取文件对象
    pub fn get_file(&self, path: &str) -> Result<File, RuntimeError> {
        if !self.file_exists(path) {
            return Err(RuntimeError::Generic(format!("文件不存在: {}", path)));
        }

        let metadata = fs::metadata(path)
            .map_err(|e| RuntimeError::Generic(format!("无法读取文件元数据: {}", e)))?;

        Ok(File {
            path: path.to_string(),
            name: Path::new(path)
                .file_name()
                .and_then(|n| n.to_str())
                .unwrap_or("")
                .to_string(),
            size: metadata.len(),
            created: metadata.created().ok().map(|t| {
                let duration_since_epoch =
                    t.duration_since(std::time::UNIX_EPOCH).unwrap_or_default();
                duration_since_epoch.as_secs() as f64
            }),
            modified: metadata.modified().ok().map(|t| {
                let duration_since_epoch =
                    t.duration_since(std::time::UNIX_EPOCH).unwrap_or_default();
                duration_since_epoch.as_secs() as f64
            }),
            accessed: metadata.accessed().ok().map(|t| {
                let duration_since_epoch =
                    t.duration_since(std::time::UNIX_EPOCH).unwrap_or_default();
                duration_since_epoch.as_secs() as f64
            }),
        })
    }

    /// 获取文件夹对象
    pub fn get_folder(&self, path: &str) -> Result<Folder, RuntimeError> {
        if !self.folder_exists(path) {
            return Err(RuntimeError::Generic(format!("文件夹不存在: {}", path)));
        }

        let _metadata = fs::metadata(path)
            .map_err(|e| RuntimeError::Generic(format!("无法读取文件夹元数据: {}", e)))?;

        // 读取文件夹内容
        let entries = fs::read_dir(path)
            .map_err(|e| RuntimeError::Generic(format!("无法读取文件夹内容: {}", e)))?;

        let mut files = Vec::new();
        let mut subfolders = Vec::new();

        for entry in entries.flatten() {
            let entry_path = entry.path();
            if entry_path.is_file() {
                if let Some(name) = entry_path.file_name().and_then(|n| n.to_str()) {
                    let size = entry.metadata().map(|m| m.len()).unwrap_or(0);
                    files.push(FileInfo {
                        name: name.to_string(),
                        size,
                    });
                }
            } else if entry_path.is_dir() {
                if let Some(name) = entry_path.file_name().and_then(|n| n.to_str()) {
                    subfolders.push(FolderInfo {
                        name: name.to_string(),
                    });
                }
            }
        }

        let name = Path::new(path)
            .file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("")
            .to_string();

        let parent = Path::new(path)
            .parent()
            .and_then(|p| p.to_str())
            .unwrap_or("")
            .to_string();

        Ok(Folder {
            path: path.to_string(),
            name,
            parent,
            _files_count: files.len() as i32,
            _subfolders_count: subfolders.len() as i32,
            files,
            subfolders,
        })
    }

    /// 获取驱动器对象
    pub fn get_drive(&self, drive_spec: &str) -> Result<Drive, RuntimeError> {
        // 简化实现
        Ok(Drive {
            letter: drive_spec.to_uppercase(),
            drive_type: "Fixed".to_string(),
            is_ready: true,
            file_system: "NTFS".to_string(),
            total_size: 107374182400.0,     // 100 GB
            available_space: 53687091200.0, // 50 GB
        })
    }
}

impl Default for FileSystemObject {
    fn default() -> Self {
        Self::new()
    }
}

impl BuiltinObject for FileSystemObject {
    fn as_any(&self) -> &dyn std::any::Any {
        self
    }

    fn as_any_mut(&mut self) -> &mut dyn std::any::Any {
        self
    }

    fn get_property(&self, _name: &str) -> Result<Value, RuntimeError> {
        Err(RuntimeError::PropertyNotFound(_name.to_string()))
    }

    fn set_property(&mut self, _name: &str, _value: Value) -> Result<(), RuntimeError> {
        Err(RuntimeError::PropertyNotFound(_name.to_string()))
    }

    fn call_method(&mut self, name: &str, args: Vec<Value>) -> Result<Value, RuntimeError> {
        match name.to_lowercase().as_str() {
            "fileexists" => {
                if args.is_empty() {
                    return Ok(Value::Boolean(false));
                }
                let path = ValueConversion::to_string(&args[0]);
                Ok(Value::Boolean(self.file_exists(&path)))
            }
            "folderexists" => {
                if args.is_empty() {
                    return Ok(Value::Boolean(false));
                }
                let path = ValueConversion::to_string(&args[0]);
                Ok(Value::Boolean(self.folder_exists(&path)))
            }
            "driveexists" => {
                if args.is_empty() {
                    return Ok(Value::Boolean(false));
                }
                let drive = ValueConversion::to_string(&args[0]);
                Ok(Value::Boolean(self.drive_exists(&drive)))
            }
            "getfile" => {
                if args.is_empty() {
                    return Err(RuntimeError::Generic("需要文件路径参数".to_string()));
                }
                let path = ValueConversion::to_string(&args[0]);
                let file = self.get_file(&path)?;
                Ok(file.to_value())
            }
            "getfolder" => {
                if args.is_empty() {
                    return Err(RuntimeError::Generic("需要文件夹路径参数".to_string()));
                }
                let path = ValueConversion::to_string(&args[0]);
                let folder = self.get_folder(&path)?;
                Ok(folder.to_value())
            }
            "getdrivename" => {
                if args.is_empty() {
                    return Ok(Value::String(String::new()));
                }
                let path = ValueConversion::to_string(&args[0]);
                Ok(Value::String(self.get_drive_name(&path)))
            }
            "getdrive" => {
                if args.is_empty() {
                    return Err(RuntimeError::Generic("需要驱动器规格参数".to_string()));
                }
                let drive_spec = ValueConversion::to_string(&args[0]);
                let drive = self.get_drive(&drive_spec)?;
                Ok(drive.to_value())
            }
            _ => Err(RuntimeError::MethodNotFound(name.to_string())),
        }
    }
}

/// 文件对象
#[derive(Debug, Clone)]
pub struct File {
    path: String,
    name: String,
    size: u64,
    created: Option<f64>,
    modified: Option<f64>,
    accessed: Option<f64>,
}

impl File {
    fn to_value(&self) -> Value {
        use std::sync::{Arc, Mutex};

        let mut dict = std::collections::HashMap::new();
        dict.insert("Path".to_string(), Value::String(self.path.clone()));
        dict.insert("Name".to_string(), Value::String(self.name.clone()));
        dict.insert("Size".to_string(), Value::Number(self.size as f64));
        dict.insert("Type".to_string(), Value::String("ASP File".to_string()));

        if let Some(created) = self.created {
            dict.insert("DateCreated".to_string(), Value::Number(created));
        }
        if let Some(modified) = self.modified {
            dict.insert("DateLastModified".to_string(), Value::Number(modified));
        }
        if let Some(accessed) = self.accessed {
            dict.insert("DateLastAccessed".to_string(), Value::Number(accessed));
        }

        // 创建一个简单的对象表示
        Value::Object(Arc::new(Mutex::new(
            crate::runtime::objects::Dictionary::from_hashmap(dict),
        )))
    }
}

/// 文件夹对象
#[derive(Debug, Clone)]
pub struct Folder {
    path: String,
    name: String,
    parent: String,
    _files_count: i32,
    _subfolders_count: i32,
    files: Vec<FileInfo>,
    subfolders: Vec<FolderInfo>,
}

#[derive(Debug, Clone)]
struct FileInfo {
    name: String,
    size: u64,
}

#[derive(Debug, Clone)]
struct FolderInfo {
    name: String,
}

impl Folder {
    fn to_value(&self) -> Value {
        use std::sync::{Arc, Mutex};

        let mut dict = std::collections::HashMap::new();
        dict.insert("Path".to_string(), Value::String(self.path.clone()));
        dict.insert("Name".to_string(), Value::String(self.name.clone()));
        dict.insert(
            "ParentFolder".to_string(),
            Value::String(self.parent.clone()),
        );
        dict.insert("ShortPath".to_string(), Value::String(self.path.clone()));
        dict.insert("ShortName".to_string(), Value::String(self.name.clone()));

        // 创建 Files 集合
        let files_array: Vec<Value> = self
            .files
            .iter()
            .map(|f| {
                let mut file_dict = std::collections::HashMap::new();
                file_dict.insert("Name".to_string(), Value::String(f.name.clone()));
                file_dict.insert("Size".to_string(), Value::Number(f.size as f64));
                Value::Object(Arc::new(Mutex::new(
                    crate::runtime::objects::Dictionary::from_hashmap(file_dict),
                )))
            })
            .collect();

        dict.insert(
            "Files".to_string(),
            Value::Array(Arc::new(Mutex::new(crate::runtime::VbsArray::from_vec(
                files_array,
            )))),
        );

        // 创建 SubFolders 集合
        let subfolders_array: Vec<Value> = self
            .subfolders
            .iter()
            .map(|f| {
                let mut folder_dict = std::collections::HashMap::new();
                folder_dict.insert("Name".to_string(), Value::String(f.name.clone()));
                Value::Object(Arc::new(Mutex::new(
                    crate::runtime::objects::Dictionary::from_hashmap(folder_dict),
                )))
            })
            .collect();

        dict.insert(
            "SubFolders".to_string(),
            Value::Array(Arc::new(Mutex::new(crate::runtime::VbsArray::from_vec(
                subfolders_array,
            )))),
        );

        Value::Object(Arc::new(Mutex::new(
            crate::runtime::objects::Dictionary::from_hashmap(dict),
        )))
    }
}

/// 驱动器对象
#[derive(Debug, Clone)]
pub struct Drive {
    letter: String,
    drive_type: String,
    is_ready: bool,
    file_system: String,
    total_size: f64,
    available_space: f64,
}

impl Drive {
    fn to_value(&self) -> Value {
        use std::sync::{Arc, Mutex};

        let mut dict = std::collections::HashMap::new();
        dict.insert(
            "DriveLetter".to_string(),
            Value::String(self.letter.clone()),
        );
        dict.insert(
            "DriveType".to_string(),
            Value::String(self.drive_type.clone()),
        );
        dict.insert("IsReady".to_string(), Value::Boolean(self.is_ready));
        dict.insert(
            "FileSystem".to_string(),
            Value::String(self.file_system.clone()),
        );
        dict.insert("TotalSize".to_string(), Value::Number(self.total_size));
        dict.insert(
            "AvailableSpace".to_string(),
            Value::Number(self.available_space),
        );

        Value::Object(Arc::new(Mutex::new(
            crate::runtime::objects::Dictionary::from_hashmap(dict),
        )))
    }
}
