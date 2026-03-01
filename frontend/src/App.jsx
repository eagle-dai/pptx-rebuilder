import React, { useState } from "react";

// 内联基础样式，确保中文字体渲染正确
const styles = {
  container: {
    fontFamily: "'PingFang SC', 'Microsoft YaHei', sans-serif",
    maxWidth: "600px",
    margin: "50px auto",
    padding: "20px",
    border: "1px solid #eee",
    borderRadius: "8px",
    boxShadow: "0 2px 10px rgba(0,0,0,0.05)",
    textAlign: "center",
  },
  button: {
    backgroundColor: "#0070f2", // 类似 SAP 的品牌蓝
    color: "white",
    border: "none",
    padding: "10px 20px",
    borderRadius: "4px",
    cursor: "pointer",
    fontSize: "16px",
    marginTop: "20px",
  },
  buttonDisabled: {
    backgroundColor: "#ccc",
    cursor: "not-allowed",
  },
  status: {
    marginTop: "20px",
    color: "#666",
  },
};

function App() {
  const [file, setFile] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [statusMsg, setStatusMsg] = useState("");

  const handleFileChange = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      setFile(e.target.files[0]);
      setStatusMsg("");
    }
  };

  const handleUpload = async () => {
    if (!file) {
      setStatusMsg("请先选择一个 PPTX 文件。");
      return;
    }

    setIsProcessing(true);
    setStatusMsg(
      "文件正在上传并由 AI 进行多模态解析重构，这可能需要几分钟时间，请稍候...",
    );

    const formData = new FormData();
    formData.append("file", file);

    try {
      const response = await fetch("http://localhost:8000/api/convert", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || "服务器返回错误");
      }

      // 处理文件下载
      const blob = await response.blob();
      const downloadUrl = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = downloadUrl;
      a.download = `Editable_${file.name}`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(downloadUrl);

      setStatusMsg("转换完成！文件已开始下载。");
    } catch (error) {
      console.error("转换出错:", error);
      setStatusMsg(`转换失败: ${error.message}`);
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div style={styles.container}>
      <h2>AI 幻灯片重构工具</h2>
      <p>
        上传由图片构成的 PPTX，使用 Claude Opus 4.5
        视觉模型提取并重构为可编辑文本的 PPTX 文件。
      </p>

      <div style={{ margin: "30px 0" }}>
        <input
          type="file"
          accept=".pptx"
          onChange={handleFileChange}
          disabled={isProcessing}
        />
      </div>

      <button
        style={
          isProcessing
            ? { ...styles.button, ...styles.buttonDisabled }
            : styles.button
        }
        onClick={handleUpload}
        disabled={isProcessing}
      >
        {isProcessing ? "AI 正在努力重构中..." : "开始转换"}
      </button>

      {statusMsg && <div style={styles.status}>{statusMsg}</div>}
    </div>
  );
}

export default App;
