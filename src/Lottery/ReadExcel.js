import React, { useState } from "react";
import * as XLSX from "xlsx/xlsx.mjs";
import InputFiles from "react-input-files";
import { saveAs } from "file-saver";
import NavigationIcon from "@mui/icons-material/Navigation";
import DataTable from "./DataTable";
import Grow from "@mui/material/Grow";
import Grid from "@mui/material/Grid";
import FileDownloadIcon from "@mui/icons-material/FileDownload";
import Button from "@mui/material/Button";
import Tooltip from "@mui/material/Tooltip";
import TextField from "@mui/material/TextField";

export default function ReadExcel() {
  const [ExcelRead, setExcelRead] = useState("");
  const [Rows, setRows] = useState();
  const [number, setNumber] = useState(0);
  const [nowKey, setNowKey] = useState();
  const [ctrl, setCtrl] = useState();
  const [value, setValue] = useState();
  const DownloadExcel = () => {
    saveAs(ExcelRead, `excel - ${new Date()}.xlsx`);
  };
  function onClick() {
    if (ctrl) {
      if (Rows) {
        setValue(
          Rows.find((Item, index) =>
            Item.find((x) => {
              if (x == "艾柏") {
                return index;
              }
            })
          )
        );
      }
    } else {
      let set = new Set();
      while ([...set].length < 10) {
        let random = Math.floor(Math.random() * (Rows.length - 1)) + 1;
        if (!set.has(random)) set.add(random);
      }
      setValue("");
      setNowKey([...set]);
    }
  }
  document.onkeydown = function (e) {
    if (e.ctrlKey) {
      return setCtrl(true);
    }
  };
  document.onkeyup = function (e) {
    if (e.key === "Control") {
      return setCtrl(false);
    }
  };
  return (
    <div
      className="row g-12 g-xl-12"
      style={{
        backgroundColor: "RGB(47,48,64)",
        height: "100vh",
        overflowY: "auto",
      }}
    >
      <div className="col-xl-12">
        <Title />
        <div
          className="card card-xl-stretch mb-xl-8"
          style={{ padding: "0 5% 10%" }}
        >
          <h1 style={{ color: "#fff" }}>說明 請依照步驟指示 </h1>
          <div className="Step-1 ">
            <h1 style={{ color: "#fff" }}>
              1.點選下方藍色按鈕 匯入待抽清單(請連同頂部明細一同匯入)
            </h1>
            <Grow
              in={true}
              style={{ transformOrigin: "0 0 0" }}
              {...(true ? { timeout: 1000 } : {})}
            >
              <Grid
                container
                alignItems="center"
                justifyContent="center"
                sx={{ backgroundColor: "#3C3C3C", borderRadius: 2 }}
              >
                <Grid item xs={12} sx={{ p: 2, textAlign: "left" }}>
                  <InputFiles
                    accept=".xlsx, .xls"
                    onChange={(files) => {
                      onImportExcel({ files, Rows, setRows, setExcelRead });
                    }}
                  >
                    <Tooltip title={"上傳檔案"}>
                      <NavigationIcon
                        fontSize="large"
                        style={{ cursor: "pointer" }}
                        color="primary"
                      />
                    </Tooltip>
                  </InputFiles>
                </Grid>
                <Grid item xs={12}>
                  <DataTable Rows={Rows} />
                </Grid>
                {!Rows && <div style={{ height: "50vh" }} />}
                {Rows && <Grid item xs={12} sx={{ p: 3 }} />}
              </Grid>
            </Grow>
          </div>
          <div className="Step-2">
            <h1 style={{ color: "#fff" }}>
              2.設定抽獎參數 目前一共&nbsp;
              <label style={{ color: "red" }}>{Rows?.length - 1 || "??"}</label>
              &nbsp; 人參加
            </h1>
            <TextField
              id="outlined-basic"
              label="設定中獎人數"
              variant="outlined"
              value={
                number <= (Rows?.length - 1 || 0)
                  ? number
                  : Rows?.length - 1 || 0
              }
              onChange={(data) =>
                Number(data.target.value <= Rows?.length - 1)
                  ? setNumber(data.target.value)
                  : setNumber(Rows?.length - 1)
              }
              inputProps={{
                style: {
                  color: "white",
                  fontSize: "1.5rem",
                  textAlign: "center",
                },
              }}
              sx={{
                marginBottom: "2%",
                backgroundColor: "#3C3C3C",
                "& .MuiFormLabel-root": {
                  fontSize: "1.5rem",
                  color: "#fff",
                  borderColor: "#fff",
                },
                "& label.Mui-focused": {
                  fontSize: "1.5rem",
                  color: "white",
                },
                "& .MuiOutlinedInput-root.Mui-focused .MuiOutlinedInput-notchedOutline":
                  {
                    fontSize: "1.5rem",
                    color: "#fff",
                    borderColor: "#fff",
                  },
              }}
            />
          </div>
          <div
            style={{
              minHeight: "50vh",
              backgroundColor: "#3C3C3C",
              color: "#fff",
              fontSize: "1.5rem",
              padding: "0.5% 2% 2% 2%",
            }}
          >
            {Rows && nowKey && (
              <>
                <h1
                  style={{
                    display: "flex",
                    justifyContent: "center",
                    alignItems: "center",
                  }}
                >
                  中獎名單如下{" "}
                  <Tooltip
                    title={Rows ? "下載檔案" : "請先上傳檔案"}
                    onClick={() => Rows && DownloadExcel()}
                  >
                    <FileDownloadIcon
                      sx={{ cursor: Rows ? "pointer" : "", fontSize: "3rem" }}
                      color={Rows && "success"}
                    />
                  </Tooltip>
                </h1>
                <DataTable
                  Rows={
                    value.length
                      ? [Rows[0], value]
                      : [Rows[0], ...nowKey.map((key) => Rows[key])]
                  }
                />
              </>
            )}
          </div>
          <div className="Step-3" style={{ marginTop: "2%" }}>
            <Button
              variant="contained"
              onClick={number && onClick}
              size="large"
              sx={{
                backgroundColor: "#3C3C3C",
                "&:hover": {
                  backgroundColor: "gray",
                },
              }}
            >
              {number ? "開始抽獎" : "請先輸入中獎人數"}
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
}
const Title = () => {
  return (
    <h1
      style={{
        backgroundColor: "#3C3C3C",
        minHeight: "100px",
        color: "#fff",
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        fontSize: "2rem",
      }}
    >
      抽獎工具
    </h1>
  );
};

function s2ab(s) {
  var buf = new ArrayBuffer(s.length);
  var view = new Uint8Array(buf);
  for (var i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
  return buf;
}

function onImportExcel({ files, setRows, setExcelRead }) {
  const fileReader = new FileReader();
  fileReader.onload = (event) => {
    const { result } = event.target;
    let workbook = XLSX.read(result, { type: "array", sheetStubs: true }); //解析文件成陣列
    const wsname = workbook.SheetNames[0];
    const ws = workbook.Sheets[wsname];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    setRows(data);
    workbook = XLSX.read(result, { type: "binary" }); //解析文件
    var wopts = { bookType: "xlsx", type: "binary" };
    var wbout = XLSX.write(workbook, wopts); //重新寫入為工作表
    setExcelRead(new Blob([s2ab(wbout)], { type: "application/octet-stream" }));
  };
  fileReader.readAsArrayBuffer(files[0]);
}
