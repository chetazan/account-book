import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import config from "./config.json";
import "./App.css";

function App() {
  const [transactions, setTransactions] = useState([]);
  const [formData, setFormData] = useState({
    type: "expense",
    category: "",
    subCategory: "",
    account: "",
    amount: "",
    description: "",
    date: new Date().toISOString().split("T")[0],
  });
  const [filter, setFilter] = useState("all");
  const [startDate, setStartDate] = useState("");
  const [endDate, setEndDate] = useState("");
  const [sortBy, setSortBy] = useState("date"); // 정렬 기준: 'date', 'account', 'category'
  const [sortOrder, setSortOrder] = useState("desc"); // 정렬 순서: 'asc', 'desc'
  // config.json에서 불러온 데이터 (하드코딩)
  const categories = config.categories;
  const subCategories = config.subCategories;
  const accounts = config.accounts;
  const [editingTransaction, setEditingTransaction] = useState(null);
  const [editingTransactionData, setEditingTransactionData] = useState({
    date: "",
    type: "expense",
    category: "",
    subCategory: "",
    account: "",
    amount: "",
    description: "",
  });
  const [showAccountModal, setShowAccountModal] = useState(false);
  const [selectedAccountForUpload, setSelectedAccountForUpload] = useState("");
  const [showFileNameModal, setShowFileNameModal] = useState(false);
  const [exportFileName, setExportFileName] = useState("");
  const fileInputRef = useRef(null);
  const mainFileInputRef = useRef(null);
  const isInitialLoad = useRef(true);

  // 로컬 스토리지에서 거래 내역 불러오기 (새로고침 시)
  useEffect(() => {
    const saved = localStorage.getItem("accountBook");
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          setTransactions(parsed);
        }
      } catch (e) {
        console.error("로컬스토리지 파싱 오류:", e);
      }
    }
    isInitialLoad.current = false;
  }, []);

  // 거래 내역 변경 시 로컬스토리지에 실시간 저장
  useEffect(() => {
    if (!isInitialLoad.current) {
      try {
        localStorage.setItem("accountBook", JSON.stringify(transactions));
        console.log("로컬스토리지 자동 저장 완료:", transactions.length, "건");
      } catch (error) {
        console.error("자동 저장 오류:", error);
      }
    }
  }, [transactions]);

  // 엑셀 파일 읽기 함수 (현재 내보내기 형식 파싱)
  const parseAccountBookExcel = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });

          // 먼저 "거래 내역" 시트 찾기 (기존 포맷)
          let transactionSheet = null;
          let sheetName = null;
          for (const name of workbook.SheetNames) {
            if (name === "거래 내역") {
              transactionSheet = workbook.Sheets[name];
              sheetName = name;
              break;
            }
          }

          // "거래 내역" 시트가 없으면 첫 번째 시트 사용 (새 포맷)
          if (!transactionSheet && workbook.SheetNames.length > 0) {
            sheetName = workbook.SheetNames[0];
            transactionSheet = workbook.Sheets[sheetName];
          }

          if (!transactionSheet) {
            reject(new Error("엑셀 파일에서 시트를 찾을 수 없습니다."));
            return;
          }

          // JSON으로 변환
          const jsonData = XLSX.utils.sheet_to_json(transactionSheet, {
            header: 1,
            defval: "",
          });

          if (jsonData.length < 2) {
            reject(new Error("거래 내역이 없습니다."));
            return;
          }

          // 헤더 행 확인
          const header = jsonData[0];

          // 입금/출금 포맷 확인 (날짜, 내역, 입금, 출금)
          const dateIndex = header.findIndex(
            (h) => String(h || "").trim() === "날짜"
          );
          const detailIndex = header.findIndex(
            (h) => String(h || "").trim() === "내역"
          );
          const depositIndex = header.findIndex(
            (h) => String(h || "").trim() === "입금"
          );
          const withdrawalIndex = header.findIndex(
            (h) => String(h || "").trim() === "출금"
          );

          // 입금/출금 포맷인 경우
          if (
            dateIndex !== -1 &&
            detailIndex !== -1 &&
            (depositIndex !== -1 || withdrawalIndex !== -1)
          ) {
            const transactions = [];

            // 데이터 행 파싱 (헤더 제외)
            for (let i = 1; i < jsonData.length; i++) {
              const row = jsonData[i];

              // 빈 행 건너뛰기
              if (!row || row.length === 0) {
                continue;
              }

              const date = String(row[dateIndex] || "").trim();
              const description = String(row[detailIndex] || "").trim();
              const deposit =
                row[depositIndex] !== undefined
                  ? parseFloat(row[depositIndex])
                  : 0;
              const withdrawal =
                row[withdrawalIndex] !== undefined
                  ? parseFloat(row[withdrawalIndex])
                  : 0;

              // 날짜와 내역이 없으면 건너뛰기
              if (!date || !description) {
                continue;
              }

              // 입금과 출금이 모두 0이거나 없으면 건너뛰기
              if (
                (!deposit || deposit === 0) &&
                (!withdrawal || withdrawal === 0)
              ) {
                continue;
              }

              // 날짜 형식 변환
              let parsedDate = date;
              if (!isNaN(parseFloat(date)) && parseFloat(date) > 25569) {
                // 엑셀 날짜 숫자 형식
                try {
                  const excelDateNum = parseFloat(date);
                  const excelDate = XLSX.SSF.parse_date_code(excelDateNum);
                  if (excelDate) {
                    const year = excelDate.y;
                    const month = String(excelDate.m).padStart(2, "0");
                    const day = String(excelDate.d).padStart(2, "0");
                    parsedDate = `${year}-${month}-${day}`;
                  }
                } catch (e) {
                  // 파싱 실패 시 원본 사용
                }
              } else if (date.includes("/") || date.includes("-")) {
                // 일반 날짜 문자열 처리
                const dateObj = new Date(date);
                if (!isNaN(dateObj.getTime())) {
                  parsedDate = dateObj.toISOString().split("T")[0];
                }
              }

              // 입금이 있으면 수입으로 추가
              if (deposit && deposit > 0) {
                transactions.push({
                  id: Date.now() + i * 2,
                  type: "income",
                  category: "",
                  subCategory: "",
                  account: "",
                  amount: deposit,
                  description: description,
                  date: parsedDate,
                });
              }

              // 출금이 있으면 지출로 추가
              if (withdrawal && withdrawal > 0) {
                transactions.push({
                  id: Date.now() + i * 2 + 1,
                  type: "expense",
                  category: "",
                  subCategory: "",
                  account: "",
                  amount: withdrawal,
                  description: description,
                  date: parsedDate,
                });
              }
            }

            if (transactions.length === 0) {
              reject(new Error("읽을 수 있는 거래 내역이 없습니다."));
              return;
            }

            resolve(transactions);
            return;
          }

          // 기존 포맷 처리 (거래 내역 시트 형식)
          const expectedHeaders = [
            "날짜",
            "유형",
            "카테고리",
            "서브카테고리",
            "계좌",
            "금액",
            "내용",
          ];

          // 헤더 인덱스 찾기
          const headerIndices = {};
          expectedHeaders.forEach((expectedHeader, idx) => {
            const foundIndex = header.findIndex((h) => h === expectedHeader);
            if (foundIndex !== -1) {
              headerIndices[expectedHeader] = foundIndex;
            }
          });

          if (Object.keys(headerIndices).length < 4) {
            reject(
              new Error(
                "엑셀 파일 형식이 올바르지 않습니다. (지원 형식: 날짜/내역/입금/출금 또는 거래 내역 시트)"
              )
            );
            return;
          }

          const transactions = [];

          // 데이터 행 파싱 (헤더 제외)
          for (let i = 1; i < jsonData.length; i++) {
            const row = jsonData[i];

            // 빈 행 건너뛰기
            if (!row || row.length === 0 || !row[headerIndices["날짜"]]) {
              continue;
            }

            const date = String(row[headerIndices["날짜"]] || "").trim();
            const type = String(row[headerIndices["유형"]] || "").trim();
            const category = String(
              row[headerIndices["카테고리"]] || ""
            ).trim();
            const subCategory = String(
              row[headerIndices["서브카테고리"]] || ""
            ).trim();
            const account = String(row[headerIndices["계좌"]] || "").trim();
            const amount = row[headerIndices["금액"]];
            const description = String(row[headerIndices["내용"]] || "").trim();

            // 필수 필드 검증
            if (
              !date ||
              !type ||
              amount === "" ||
              amount === null ||
              amount === undefined
            ) {
              continue;
            }

            // 날짜 형식 변환 (엑셀 날짜 숫자 형식 처리)
            let parsedDate = date;
            if (!isNaN(parseFloat(date)) && parseFloat(date) > 25569) {
              // 엑셀 날짜 숫자 형식
              try {
                const excelDateNum = parseFloat(date);
                const excelDate = XLSX.SSF.parse_date_code(excelDateNum);
                if (excelDate) {
                  const year = excelDate.y;
                  const month = String(excelDate.m).padStart(2, "0");
                  const day = String(excelDate.d).padStart(2, "0");
                  parsedDate = `${year}-${month}-${day}`;
                }
              } catch (e) {
                // 파싱 실패 시 원본 사용
              }
            } else if (date.includes("/") || date.includes("-")) {
              // 일반 날짜 문자열 처리
              const dateObj = new Date(date);
              if (!isNaN(dateObj.getTime())) {
                parsedDate = dateObj.toISOString().split("T")[0];
              }
            }

            // 금액 처리
            let parsedAmount = parseFloat(amount);
            if (isNaN(parsedAmount) || parsedAmount <= 0) {
              continue;
            }

            // 유형 변환
            const transactionType = type === "수입" ? "income" : "expense";

            transactions.push({
              id: Date.now() + i,
              type: transactionType,
              category: category || "",
              subCategory: subCategory || "",
              account: account || "",
              amount: parsedAmount,
              description: description || "",
              date: parsedDate,
            });
          }

          if (transactions.length === 0) {
            reject(new Error("읽을 수 있는 거래 내역이 없습니다."));
            return;
          }

          resolve(transactions);
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = (error) => {
        console.error("파일 읽기 오류:", error);
        reject(new Error("파일을 읽을 수 없습니다."));
      };

      reader.readAsArrayBuffer(file);
    });
  };

  // 메인 엑셀 파일 로드 핸들러
  const handleLoadMainFile = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
      alert("엑셀 파일(.xlsx, .xls)만 지원합니다.");
      if (mainFileInputRef.current) {
        mainFileInputRef.current.value = "";
      }
      return;
    }

    try {
      const transactions = await parseAccountBookExcel(file);

      if (transactions.length > 0) {
        setTransactions(transactions);
        // 로컬스토리지에 저장 (useEffect에서 자동 저장되지만 명시적으로 저장)
        localStorage.setItem("accountBook", JSON.stringify(transactions));
        alert(`${transactions.length}개의 거래 내역을 불러왔습니다.`);
      } else {
        alert("거래 내역이 없습니다.");
      }
    } catch (error) {
      console.error("엑셀 파일 읽기 오류:", error);
      alert("엑셀 파일을 읽는 중 오류가 발생했습니다:\n\n" + error.message);
    }

    // 파일 입력 초기화
    if (mainFileInputRef.current) {
      mainFileInputRef.current.value = "";
    }
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!formData.category || !formData.amount || !formData.description) {
      alert("모든 필드를 입력해주세요.");
      return;
    }

    const newTransaction = {
      id: Date.now(),
      ...formData,
      amount: parseFloat(formData.amount),
    };

    setTransactions([newTransaction, ...transactions]);
    setFormData({
      type: "expense",
      category: "",
      subCategory: "",
      account: config.accounts.length > 0 ? config.accounts[0] : "",
      amount: "",
      description: "",
      date: new Date().toISOString().split("T")[0],
    });
  };

  // 계좌 초기화 (formData.account 업데이트)
  useEffect(() => {
    if (
      config.accounts.length > 0 &&
      (!formData.account || !config.accounts.includes(formData.account))
    ) {
      setFormData((prev) => ({ ...prev, account: config.accounts[0] }));
    }
  }, [formData.account]);

  const handleDelete = (id) => {
    if (window.confirm("정말 삭제하시겠습니까?")) {
      setTransactions(transactions.filter((t) => t.id !== id));
    }
  };

  const handleDeleteAll = () => {
    if (
      window.confirm(
        "모든 거래 내역을 삭제하시겠습니까?\n\n이 작업은 되돌릴 수 없습니다."
      )
    ) {
      setTransactions([]);
      localStorage.setItem("accountBook", JSON.stringify([]));
    }
  };

  const handleEditTransaction = (transaction) => {
    setEditingTransaction(transaction.id);
    setEditingTransactionData({
      date: transaction.date || "",
      type: transaction.type || "expense",
      category: transaction.category || "",
      subCategory: transaction.subCategory || "",
      account: transaction.account || "",
      amount: transaction.amount ? transaction.amount.toString() : "",
      description: transaction.description || "",
    });
  };

  const handleSaveTransaction = (id) => {
    const transaction = transactions.find((t) => t.id === id);
    if (!transaction) return;

    // 필수 필드 검증
    if (
      !editingTransactionData.date ||
      !editingTransactionData.category ||
      !editingTransactionData.amount ||
      !editingTransactionData.description
    ) {
      alert("모든 필수 필드를 입력해주세요.");
      return;
    }

    // 계좌 유효성 검사
    if (
      !editingTransactionData.account ||
      !config.accounts.includes(editingTransactionData.account)
    ) {
      alert("유효한 계좌를 선택해주세요.");
      return;
    }

    // 카테고리 유효성 검사
    const validCategories =
      config.categories[editingTransactionData.type] || [];
    if (!validCategories.includes(editingTransactionData.category)) {
      alert("유효한 카테고리를 선택해주세요.");
      return;
    }

    // 서브카테고리 유효성 검사
    if (editingTransactionData.subCategory) {
      const validSubCategories =
        config.subCategories[editingTransactionData.category] || [];
      if (!validSubCategories.includes(editingTransactionData.subCategory)) {
        alert("유효한 서브 카테고리를 선택해주세요.");
        return;
      }
    }

    // 금액 검증
    const amount = parseFloat(editingTransactionData.amount);
    if (isNaN(amount) || amount <= 0) {
      alert("유효한 금액을 입력해주세요.");
      return;
    }

    setTransactions(
      transactions.map((t) =>
        t.id === id
          ? {
              ...t,
              date: editingTransactionData.date,
              type: editingTransactionData.type,
              category: editingTransactionData.category,
              subCategory: editingTransactionData.subCategory,
              account: editingTransactionData.account,
              amount: amount,
              description: editingTransactionData.description,
            }
          : t
      )
    );
    setEditingTransaction(null);
    setEditingTransactionData({
      date: "",
      type: "expense",
      category: "",
      subCategory: "",
      account: "",
      amount: "",
      description: "",
    });
  };

  const handleCancelEdit = () => {
    setEditingTransaction(null);
    setEditingTransactionData({
      date: "",
      type: "expense",
      category: "",
      subCategory: "",
      account: "",
      amount: "",
      description: "",
    });
  };

  // 날짜 형식 변환: "2025년 10월 01일" -> "2025-10-01"
  const parseKoreanDate = (dateStr) => {
    if (!dateStr || dateStr === "-" || dateStr.trim() === "") return null;

    const match = dateStr.match(/(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일/);
    if (match) {
      const year = match[1];
      const month = match[2].padStart(2, "0");
      const day = match[3].padStart(2, "0");
      return `${year}-${month}-${day}`;
    }
    return null;
  };

  // 가맹점명으로 카테고리 자동 분류
  const categorizeByMerchant = (merchant) => {
    if (!merchant) return "기타 지출";

    const merchantLower = merchant.toLowerCase();

    // 식비
    if (
      merchant.includes("카페") ||
      merchant.includes("커피") ||
      merchant.includes("베이커리") ||
      merchant.includes("식당") ||
      merchant.includes("호프") ||
      merchant.includes("미역") ||
      merchant.includes("마트") ||
      merchant.includes("편의점") ||
      merchant.includes("CU") ||
      merchant.includes("GS") ||
      merchant.includes("세븐일레븐")
    ) {
      return "식비";
    }

    // 교통비
    if (
      merchant.includes("티머니") ||
      merchant.includes("지하철") ||
      merchant.includes("버스") ||
      merchant.includes("휴게소") ||
      merchant.includes("주유소")
    ) {
      return "교통비";
    }

    // 쇼핑
    if (
      merchant.includes("쿠팡") ||
      merchant.includes("네이버") ||
      merchant.includes("쇼핑") ||
      merchant.includes("다이소") ||
      merchant.includes("마켓")
    ) {
      return "쇼핑";
    }

    // 문화생활
    if (
      merchant.includes("영화") ||
      merchant.includes("공연") ||
      merchant.includes("레저")
    ) {
      return "문화생활";
    }

    // 의료비
    if (
      merchant.includes("병원") ||
      merchant.includes("약국") ||
      merchant.includes("의료")
    ) {
      return "의료비";
    }

    // 주거비
    if (
      merchant.includes("관리비") ||
      merchant.includes("전기") ||
      merchant.includes("가스") ||
      merchant.includes("수도") ||
      merchant.includes("월세") ||
      merchant.includes("임대")
    ) {
      return "주거비";
    }

    return "기타 지출";
  };

  // HTML 파일 파싱 (현대카드 명세서 형식)
  const parseHTMLFile = (htmlContent, selectedAccount = "") => {
    try {
      console.log("HTML 파싱 시작, 파일 크기:", htmlContent.length);
      const parser = new DOMParser();
      const doc = parser.parseFromString(htmlContent, "text/html");

      // 모든 테이블 찾기
      const tables = doc.querySelectorAll("table");
      console.log("찾은 테이블 수:", tables.length);

      if (tables.length === 0) {
        console.log("테이블을 찾을 수 없습니다.");
        return [];
      }

      // 첫 번째 테이블 사용
      const table = tables[0];
      const rows = table.querySelectorAll("tr");
      console.log("테이블 행 수:", rows.length);

      const transactions = [];

      // 헤더 행 찾기 (이용일, 이용가맹점, 이용금액이 있는 행)
      let headerRowIndex = -1;
      for (let i = 0; i < rows.length; i++) {
        const cells = rows[i].querySelectorAll("th, td");
        if (cells.length === 0) continue;

        const cellTexts = Array.from(cells).map((cell) => {
          const text = cell.textContent.trim();
          console.log(`행 ${i}, 셀 텍스트:`, text);
          return text;
        });

        // 더 유연한 검색: 일부 키워드만 있어도 인식
        const hasDate = cellTexts.some(
          (text) =>
            text.includes("이용일") ||
            text.includes("날짜") ||
            text.includes("Date")
        );
        const hasMerchant = cellTexts.some(
          (text) =>
            text.includes("가맹점") ||
            text.includes("내용") ||
            text.includes("Merchant")
        );
        const hasAmount = cellTexts.some(
          (text) =>
            text.includes("이용금액") ||
            text.includes("금액") ||
            text.includes("Amount")
        );

        console.log(`행 ${i} 검사:`, {
          hasDate,
          hasMerchant,
          hasAmount,
          cellTexts: cellTexts.slice(0, 5),
        });

        if (hasDate && hasMerchant && hasAmount) {
          headerRowIndex = i;
          console.log("헤더 행 발견:", i, cellTexts);
          break;
        }
      }

      if (headerRowIndex === -1) {
        console.log("헤더 행을 찾을 수 없습니다. 모든 행의 첫 5개 셀:");
        for (let i = 0; i < Math.min(10, rows.length); i++) {
          const cells = rows[i].querySelectorAll("th, td");
          const cellTexts = Array.from(cells)
            .slice(0, 5)
            .map((cell) => cell.textContent.trim());
          console.log(`행 ${i}:`, cellTexts);
        }
        return [];
      }

      // 헤더 행에서 컬럼 인덱스 찾기
      const headerCells = rows[headerRowIndex].querySelectorAll("th, td");
      let dateIndex = -1;
      let merchantIndex = -1;
      let amountIndex = -1;
      let feeIndex = -1; // 수수료 컬럼

      headerCells.forEach((cell, index) => {
        const text = cell.textContent.trim();
        console.log(`헤더 셀 ${index}:`, text);
        if (
          text.includes("이용일") ||
          text.includes("날짜") ||
          text.includes("Date")
        ) {
          dateIndex = index;
          console.log("날짜 인덱스:", dateIndex);
        }
        if (
          text.includes("가맹점") ||
          text.includes("내용") ||
          text.includes("Merchant")
        ) {
          merchantIndex = index;
          console.log("가맹점 인덱스:", merchantIndex);
        }
        if (
          text.includes("이용금액") ||
          text.includes("금액") ||
          text.includes("Amount")
        ) {
          amountIndex = index;
          console.log("금액 인덱스:", amountIndex);
        }
        if (text.includes("수수료") || text.includes("이자")) {
          feeIndex = index;
          console.log("수수료 인덱스:", feeIndex);
        }
      });

      console.log("컬럼 인덱스:", {
        dateIndex,
        merchantIndex,
        amountIndex,
        feeIndex,
      });

      if (dateIndex === -1 || merchantIndex === -1 || amountIndex === -1) {
        console.log("필수 컬럼을 찾을 수 없습니다.", {
          dateIndex,
          merchantIndex,
          amountIndex,
        });
        return [];
      }

      // 데이터 행 파싱
      let parsedCount = 0;
      let skippedCount = 0;

      for (let i = headerRowIndex + 1; i < rows.length; i++) {
        const cells = rows[i].querySelectorAll("td");
        if (cells.length === 0) {
          skippedCount++;
          continue;
        }

        // 인덱스 범위 체크
        if (
          dateIndex >= cells.length ||
          merchantIndex >= cells.length ||
          amountIndex >= cells.length
        ) {
          console.log(`행 ${i}: 셀 수 부족 (${cells.length}개)`);
          skippedCount++;
          continue;
        }

        const dateCell = cells[dateIndex]?.textContent.trim() || "";
        const merchantCell = cells[merchantIndex]?.textContent.trim() || "";
        const amountCell = cells[amountIndex]?.textContent.trim() || "";
        const feeCell =
          feeIndex >= 0 && feeIndex < cells.length
            ? cells[feeIndex]?.textContent.trim() || ""
            : "";

        // 처음 몇 개 행만 상세 로그
        if (i < headerRowIndex + 5) {
          console.log(`행 ${i} 데이터:`, {
            dateCell,
            merchantCell,
            amountCell,
            feeCell,
          });
        }

        // 소계, 합계, 빈 행 건너뛰기
        if (
          !dateCell ||
          dateCell === "-" ||
          merchantCell.includes("소계") ||
          merchantCell.includes("합계") ||
          (merchantCell.includes("건") && !merchantCell.match(/\d+건/))
        ) {
          skippedCount++;
          continue;
        }

        const date = parseKoreanDate(dateCell);
        if (!date) {
          // 날짜 파싱 실패한 경우 건너뛰기
          if (i < headerRowIndex + 5) {
            console.log(`행 ${i}: 날짜 파싱 실패 - "${dateCell}"`);
          }
          skippedCount++;
          continue;
        }

        // 금액 처리: 이용금액이 0이면 수수료 컬럼 확인
        let amount = parseFloat(amountCell.replace(/,/g, ""));
        if (isNaN(amount) || amount === 0) {
          // 수수료 컬럼이 있으면 그것을 사용
          if (feeCell) {
            amount = parseFloat(feeCell.replace(/,/g, ""));
          }
          // 여전히 0이거나 유효하지 않으면 건너뛰기
          if (isNaN(amount) || amount <= 0) {
            skippedCount++;
            continue;
          }
        }

        const category = categorizeByMerchant(merchantCell);

        transactions.push({
          id: Date.now() + i,
          type: "expense",
          category: category,
          amount: amount,
          description: merchantCell,
          date: date,
          account:
            selectedAccount ||
            (config.accounts.length > 0 ? config.accounts[0] : ""),
        });
        parsedCount++;
      }

      console.log(
        `파싱 완료: ${transactions.length}개의 내역을 찾았습니다. (건너뛴 행: ${skippedCount}개)`
      );
      return transactions;
    } catch (error) {
      console.error("HTML 파싱 오류:", error);
      return [];
    }
  };

  // 엑셀 파일 파싱
  const parseExcelFile = (file, selectedAccount = "") => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: "array" });

          // 첫 번째 시트 사용
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];

          // JSON으로 변환
          const jsonData = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            defval: "",
          });

          console.log("엑셀 데이터 행 수:", jsonData.length);
          console.log("첫 10개 행:", jsonData.slice(0, 10));

          // 헤더 찾기
          let headerRowIndex = -1;
          let dateIndex = -1;
          let merchantIndex = -1;
          let summaryIndex = -1; // 적요 인덱스 (우리은행)
          let amountIndex = -1;
          let withdrawalIndex = -1; // 출금액 인덱스 (우리은행/국민은행)
          let depositIndex = -1; // 입금액 인덱스 (우리은행/국민은행)
          let senderReceiverIndex = -1; // 보낸분/받는분 인덱스 (국민은행)
          let benefitAmountIndex = -1; // 혜택금액 인덱스 (삼성카드)
          let principalIndex = -1; // 원금 인덱스 (삼성카드)
          let isWooriBankFormat = false; // 우리은행 형식 여부
          let isKbBankFormat = false; // 국민은행 형식 여부
          let isSamsungCardFormat = false; // 삼성카드 형식 여부
          let isSimpleDepositWithdrawalFormat = false; // 간단한 입금/출금 포맷 (날짜, 내역, 입금, 출금)

          for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length === 0) continue;

            // 셀 값을 문자열로 변환 (소문자 변환 없이 원본 유지)
            const rowStrings = row.map((cell) => String(cell || "").trim());
            const rowLower = rowStrings.map((cell) => cell.toLowerCase());

            // 간단한 입금/출금 포맷 감지 (날짜, 내역, 입금, 출금)
            const hasSimpleDate = rowStrings.some((cell) => {
              const cellTrimmed = cell.trim();
              return cellTrimmed === "날짜";
            });
            const hasSimpleDetail = rowStrings.some((cell) => {
              const cellTrimmed = cell.trim();
              return cellTrimmed === "내역";
            });
            const hasSimpleDeposit = rowStrings.some((cell) => {
              const cellTrimmed = cell.trim();
              return cellTrimmed === "입금";
            });
            const hasSimpleWithdrawal = rowStrings.some((cell) => {
              const cellTrimmed = cell.trim();
              return cellTrimmed === "출금";
            });

            // 간단한 입금/출금 포맷 감지 (우선순위 높음)
            if (
              hasSimpleDate &&
              hasSimpleDetail &&
              (hasSimpleDeposit || hasSimpleWithdrawal)
            ) {
              isSimpleDepositWithdrawalFormat = true;
              headerRowIndex = i;
              console.log("간단한 입금/출금 포맷 감지됨, 헤더 행:", i);
              console.log("헤더 행 데이터:", rowStrings);
              rowStrings.forEach((cellValue, idx) => {
                const cellTrimmed = cellValue.trim();
                if (cellTrimmed === "날짜") {
                  dateIndex = idx;
                  console.log(`    → 날짜 인덱스: ${idx}`);
                }
                if (cellTrimmed === "내역") {
                  merchantIndex = idx;
                  console.log(`    → 내역 인덱스: ${idx}`);
                }
                if (cellTrimmed === "입금") {
                  depositIndex = idx;
                  console.log(`    → 입금 인덱스: ${idx}`);
                }
                if (cellTrimmed === "출금") {
                  withdrawalIndex = idx;
                  console.log(`    → 출금 인덱스: ${idx}`);
                }
              });
              console.log("간단한 입금/출금 포맷 컬럼 인덱스:", {
                dateIndex,
                merchantIndex,
                depositIndex,
                withdrawalIndex,
              });
              break;
            }

            // 국민은행 형식 감지 (출금액/입금액 컬럼이 있는지 확인)
            // 국민은행은 "출금액"과 "입금액" 컬럼이 명확하게 있음
            const hasKbWithdrawal = rowStrings.some((cell) => {
              const cellTrimmed = cell.trim();
              return (
                cellTrimmed === "출금액" ||
                (cellTrimmed.includes("출금액") &&
                  !cellTrimmed.includes("가능") &&
                  !cellTrimmed.includes("출금가능"))
              );
            });
            const hasKbDeposit = rowStrings.some((cell) => {
              const cellTrimmed = cell.trim();
              return cellTrimmed === "입금액" || cellTrimmed.includes("입금액");
            });
            const hasKbSenderReceiver = rowStrings.some((cell) => {
              const cellTrimmed = cell.trim();
              return (
                cellTrimmed.includes("보낸분") ||
                cellTrimmed.includes("받는분") ||
                cellTrimmed.includes("보낸분/받는분")
              );
            });
            const hasKbDate = rowStrings.some((cell) => {
              const cellTrimmed = cell.trim();
              return (
                cellTrimmed.includes("거래일시") ||
                cellTrimmed.includes("거래일자") ||
                cellTrimmed.includes("거래일")
              );
            });

            // 우리은행 형식 감지 (찾으신금액/맡기신금액 또는 출금/입금 컬럼이 있는지 확인)
            const hasWooriWithdrawal = rowStrings.some((cell) => {
              const cellTrimmed = cell.trim();
              return (
                cellTrimmed.includes("찾으신금액") ||
                (cellTrimmed.includes("출금") &&
                  (cellTrimmed.includes("원화") || cellTrimmed.includes("원")))
              );
            });
            const hasWooriDeposit = rowStrings.some((cell) => {
              const cellTrimmed = cell.trim();
              return (
                cellTrimmed.includes("맡기신금액") ||
                (cellTrimmed.includes("입금") &&
                  (cellTrimmed.includes("원화") || cellTrimmed.includes("원")))
              );
            });

            // 삼성카드 형식 감지 (이용일, 가맹점, 이용금액, 원금 컬럼이 있는지 확인)
            const hasSamsungDate = rowStrings.some((cell) => {
              const cellTrimmed = cell.trim();
              return cellTrimmed === "이용일" || cellTrimmed.includes("이용일");
            });
            const hasSamsungMerchant = rowStrings.some((cell) => {
              const cellTrimmed = cell.trim();
              return cellTrimmed === "가맹점" || cellTrimmed.includes("가맹점");
            });
            const hasSamsungAmount = rowStrings.some((cell) => {
              const cellTrimmed = cell.trim();
              return (
                cellTrimmed === "이용금액" || cellTrimmed.includes("이용금액")
              );
            });
            const hasSamsungPrincipal = rowStrings.some((cell) => {
              const cellTrimmed = cell.trim();
              return cellTrimmed === "원금" || cellTrimmed.includes("원금");
            });

            // 디버깅: 형식 감지 시도 (처음 10개 행만)
            if (i < 10) {
              console.log(`행 ${i} 검사:`, {
                hasKbWithdrawal,
                hasKbDeposit,
                hasKbDate,
                hasSamsungDate,
                hasSamsungMerchant,
                hasSamsungAmount,
                hasSamsungPrincipal,
                rowStrings: rowStrings.slice(0, 10),
              });
            }

            // 삼성카드 형식: 이용일, 가맹점, 이용금액, 원금이 모두 있으면 삼성카드
            if (
              hasSamsungDate &&
              hasSamsungMerchant &&
              hasSamsungAmount &&
              hasSamsungPrincipal
            ) {
              isSamsungCardFormat = true;
              headerRowIndex = i;
              console.log("삼성카드 형식 감지됨, 헤더 행:", i);
              console.log("헤더 행 데이터:", rowStrings);
              rowStrings.forEach((cellValue, idx) => {
                const cellLower = cellValue.toLowerCase();
                console.log(`  컬럼 ${idx}: "${cellValue}"`);
                // 날짜 컬럼: 이용일
                if (cellValue === "이용일" || cellLower.includes("이용일")) {
                  if (dateIndex === -1) {
                    dateIndex = idx;
                    console.log(`    → 이용일 인덱스: ${idx}`);
                  }
                }
                // 가맹점 컬럼
                if (cellValue === "가맹점" || cellLower.includes("가맹점")) {
                  if (merchantIndex === -1) {
                    merchantIndex = idx;
                    console.log(`    → 가맹점 인덱스: ${idx}`);
                  }
                }
                // 이용금액 컬럼
                if (
                  cellValue === "이용금액" ||
                  cellLower.includes("이용금액")
                ) {
                  if (amountIndex === -1) {
                    amountIndex = idx;
                    console.log(`    → 이용금액 인덱스: ${idx}`);
                  }
                }
                // 혜택금액 컬럼
                if (
                  cellValue === "혜택금액" ||
                  cellLower.includes("혜택금액")
                ) {
                  if (benefitAmountIndex === -1) {
                    benefitAmountIndex = idx;
                    console.log(`    → 혜택금액 인덱스: ${idx}`);
                  }
                }
                // 원금 컬럼
                if (cellValue === "원금" || cellLower.includes("원금")) {
                  if (principalIndex === -1) {
                    principalIndex = idx;
                    console.log(`    → 원금 인덱스: ${idx}`);
                  }
                }
              });
              console.log("삼성카드 컬럼 인덱스:", {
                dateIndex,
                merchantIndex,
                amountIndex,
                benefitAmountIndex,
                principalIndex,
              });
              break;
            }
            // 국민은행 형식: 출금액과 입금액이 모두 있으면 국민은행
            else if (hasKbWithdrawal && hasKbDeposit) {
              isKbBankFormat = true;
              headerRowIndex = i;
              console.log("국민은행 형식 감지됨, 헤더 행:", i);
              console.log("헤더 행 데이터:", rowStrings);
              rowStrings.forEach((cellValue, idx) => {
                const cellLower = cellValue.toLowerCase();
                console.log(`  컬럼 ${idx}: "${cellValue}"`);
                // 날짜 컬럼: 거래일시 (정확한 매칭 우선)
                if (
                  cellValue === "거래일시" ||
                  cellValue === "거래일자" ||
                  cellLower.includes("거래일시") ||
                  cellLower.includes("거래일자") ||
                  cellLower.includes("거래일") ||
                  cellLower.includes("이용일") ||
                  cellLower.includes("날짜") ||
                  cellLower.includes("date")
                ) {
                  if (dateIndex === -1) {
                    dateIndex = idx;
                    console.log(`    → 날짜 인덱스: ${idx}`);
                  }
                }
                // 적요 컬럼
                if (cellValue === "적요" || cellLower.includes("적요")) {
                  if (summaryIndex === -1) {
                    summaryIndex = idx;
                    if (merchantIndex === -1) {
                      merchantIndex = idx;
                    }
                    console.log(`    → 적요 인덱스: ${idx}`);
                  }
                }
                // 보낸분/받는분 컬럼
                if (
                  cellValue.includes("보낸분") ||
                  cellValue.includes("받는분") ||
                  cellValue.includes("보낸분/받는분")
                ) {
                  if (senderReceiverIndex === -1) {
                    senderReceiverIndex = idx;
                    console.log(`    → 보낸분/받는분 인덱스: ${idx}`);
                  }
                }
                // 출금액 컬럼 (출금가능금액 제외)
                if (
                  cellValue === "출금액" ||
                  (cellLower.includes("출금액") &&
                    !cellLower.includes("가능") &&
                    !cellLower.includes("출금가능"))
                ) {
                  if (withdrawalIndex === -1) {
                    withdrawalIndex = idx;
                    console.log(`    → 출금액 인덱스: ${idx}`);
                  }
                }
                // 입금액 컬럼
                if (cellValue === "입금액" || cellLower.includes("입금액")) {
                  if (depositIndex === -1) {
                    depositIndex = idx;
                    console.log(`    → 입금액 인덱스: ${idx}`);
                  }
                }
              });
              console.log("국민은행 컬럼 인덱스:", {
                dateIndex,
                merchantIndex,
                summaryIndex,
                senderReceiverIndex,
                withdrawalIndex,
                depositIndex,
              });
              break;
            } else if (hasWooriWithdrawal && hasWooriDeposit) {
              isWooriBankFormat = true;
              headerRowIndex = i;
              row.forEach((cell, idx) => {
                const cellLower = String(cell).toLowerCase();
                // 날짜 컬럼: 거래일시, 거래일자, 거래일, 이용일, 날짜
                if (
                  cellLower.includes("거래일시") ||
                  cellLower.includes("거래일자") ||
                  cellLower.includes("거래일") ||
                  cellLower.includes("이용일") ||
                  cellLower.includes("날짜") ||
                  cellLower.includes("date")
                ) {
                  dateIndex = idx;
                }
                // 내용 컬럼: 기재내용을 우선, 적요는 별도로 저장
                if (cellLower.includes("기재내용")) {
                  merchantIndex = idx;
                } else if (cellLower.includes("적요")) {
                  summaryIndex = idx;
                } else if (
                  merchantIndex === -1 &&
                  (cellLower.includes("거래내용") ||
                    cellLower.includes("내용") ||
                    cellLower.includes("가맹점") ||
                    cellLower.includes("description") ||
                    cellLower.includes("merchant"))
                ) {
                  merchantIndex = idx;
                }
                // 출금 컬럼: 찾으신금액 또는 출금(원화)
                if (cellLower.includes("찾으신금액")) {
                  withdrawalIndex = idx;
                } else if (
                  withdrawalIndex === -1 &&
                  cellLower.includes("출금") &&
                  (cellLower.includes("원화") || cellLower.includes("원"))
                ) {
                  withdrawalIndex = idx;
                }
                // 입금 컬럼: 맡기신금액 또는 입금(원화)
                if (cellLower.includes("맡기신금액")) {
                  depositIndex = idx;
                } else if (
                  depositIndex === -1 &&
                  cellLower.includes("입금") &&
                  (cellLower.includes("원화") || cellLower.includes("원"))
                ) {
                  depositIndex = idx;
                }
              });
              break;
            } else if (
              rowLower.some(
                (cell) =>
                  cell.includes("이용일") ||
                  cell.includes("날짜") ||
                  cell.includes("date")
              )
            ) {
              headerRowIndex = i;
              row.forEach((cell, idx) => {
                const cellLower = String(cell).toLowerCase();
                if (
                  cellLower.includes("이용일") ||
                  cellLower.includes("날짜") ||
                  cellLower.includes("date")
                ) {
                  dateIndex = idx;
                }
                if (
                  cellLower.includes("가맹점") ||
                  cellLower.includes("내용") ||
                  cellLower.includes("description") ||
                  cellLower.includes("merchant")
                ) {
                  merchantIndex = idx;
                }
                if (
                  cellLower.includes("이용금액") ||
                  cellLower.includes("금액") ||
                  cellLower.includes("amount")
                ) {
                  amountIndex = idx;
                }
              });
              break;
            }
          }

          if (headerRowIndex === -1) {
            console.error("헤더 행을 찾을 수 없습니다.");
            console.log("엑셀 파일의 모든 행 (최대 20개):");
            jsonData.slice(0, 20).forEach((row, idx) => {
              const rowStr = row
                .map((cell) => String(cell || "").trim())
                .join(" | ");
              console.log(`행 ${idx}:`, rowStr);
            });
            reject(
              new Error(
                "엑셀 파일에서 헤더 행을 찾을 수 없습니다.\n\n파일 형식을 확인해주세요.\n(국민은행, 우리은행 거래내역 엑셀 파일 또는 일반 엑셀 형식)\n\n브라우저 개발자 도구(F12)의 콘솔 탭에서 파일 구조를 확인할 수 있습니다."
              )
            );
            return;
          }

          console.log("헤더 행 발견:", headerRowIndex);
          console.log("컬럼 인덱스:", {
            dateIndex,
            merchantIndex,
            summaryIndex,
            senderReceiverIndex,
            withdrawalIndex,
            depositIndex,
            benefitAmountIndex,
            principalIndex,
            isKbBankFormat,
            isWooriBankFormat,
            isSamsungCardFormat,
            isSimpleDepositWithdrawalFormat,
          });

          const transactions = [];

          for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i];

            // 간단한 입금/출금 포맷 처리
            if (isSimpleDepositWithdrawalFormat) {
              if (
                dateIndex === -1 ||
                merchantIndex === -1 ||
                (depositIndex === -1 && withdrawalIndex === -1)
              ) {
                continue;
              }

              const dateStr = String(row[dateIndex] || "").trim();
              const description = String(row[merchantIndex] || "").trim();
              const deposit =
                depositIndex >= 0
                  ? parseFloat(
                      String(row[depositIndex] || "").replace(/,/g, "")
                    )
                  : 0;
              const withdrawal =
                withdrawalIndex >= 0
                  ? parseFloat(
                      String(row[withdrawalIndex] || "").replace(/,/g, "")
                    )
                  : 0;

              // 날짜와 내역이 없으면 건너뛰기
              if (!dateStr || !description) {
                continue;
              }

              // 입금과 출금이 모두 0이거나 없으면 건너뛰기
              if (
                (!deposit || deposit === 0) &&
                (!withdrawal || withdrawal === 0)
              ) {
                continue;
              }

              // 날짜 형식 변환
              let date = null;
              if (!isNaN(parseFloat(dateStr)) && parseFloat(dateStr) > 25569) {
                // 엑셀 날짜 숫자 형식
                try {
                  const excelDateNum = parseFloat(dateStr);
                  const excelDate = XLSX.SSF.parse_date_code(excelDateNum);
                  if (excelDate) {
                    const year = excelDate.y;
                    const month = String(excelDate.m).padStart(2, "0");
                    const day = String(excelDate.d).padStart(2, "0");
                    date = `${year}-${month}-${day}`;
                  }
                } catch (e) {
                  // 파싱 실패 시 다음 방법 시도
                }
              } else if (dateStr.includes("/") || dateStr.includes("-")) {
                // 일반 날짜 문자열 처리
                const dateObj = new Date(dateStr);
                if (!isNaN(dateObj.getTime())) {
                  date = dateObj.toISOString().split("T")[0];
                }
              }

              if (!date) {
                continue;
              }

              // 입금이 있으면 수입으로 추가
              if (deposit && deposit > 0) {
                transactions.push({
                  id: Date.now() + i * 2,
                  type: "income",
                  category: "",
                  subCategory: "",
                  account: selectedAccount || "",
                  amount: deposit,
                  description: description,
                  date: date,
                });
              }

              // 출금이 있으면 지출로 추가
              if (withdrawal && withdrawal > 0) {
                transactions.push({
                  id: Date.now() + i * 2 + 1,
                  type: "expense",
                  category: "",
                  subCategory: "",
                  account: selectedAccount || "",
                  amount: withdrawal,
                  description: description,
                  date: date,
                });
              }

              continue;
            }

            // 삼성카드 형식 처리
            if (isSamsungCardFormat) {
              if (
                dateIndex === -1 ||
                merchantIndex === -1 ||
                principalIndex === -1
              ) {
                if (i < headerRowIndex + 5) {
                  console.log(`행 ${i}: 필수 컬럼 없음`, {
                    dateIndex,
                    merchantIndex,
                    principalIndex,
                  });
                }
                continue;
              }

              const dateStr = String(row[dateIndex] || "").trim();
              const merchantStr = String(row[merchantIndex] || "").trim();
              const principalStr =
                principalIndex >= 0
                  ? String(row[principalIndex] || "").trim()
                  : "";

              if (i < headerRowIndex + 5) {
                console.log(`행 ${i} 데이터:`, {
                  dateStr,
                  merchantStr,
                  principalStr,
                });
              }

              // 합계 행 건너뛰기
              const isSummaryRow = row.some((cell) => {
                const cellStr = String(cell || "").trim();
                return (
                  cellStr === "합계" ||
                  cellStr === "소계" ||
                  cellStr.includes("합계") ||
                  cellStr.includes("소계") ||
                  cellStr.includes("일시불합계")
                );
              });

              // 소계, 합계, 빈 날짜 건너뛰기
              if (
                !dateStr ||
                dateStr === "-" ||
                isSummaryRow ||
                merchantStr.includes("소계") ||
                merchantStr.includes("합계")
              ) {
                if (i < headerRowIndex + 5) {
                  console.log(`행 ${i}: 건너뛰기 (합계/소계 행 또는 빈 날짜)`);
                }
                continue;
              }

              // 원금이 없거나 0이면 건너뛰기
              const principalAmount = principalStr
                ? parseFloat(principalStr.toString().replace(/,/g, ""))
                : 0;
              if (
                !principalStr ||
                isNaN(principalAmount) ||
                principalAmount === 0
              ) {
                if (i < headerRowIndex + 5) {
                  console.log(`행 ${i}: 건너뛰기 (원금 없음)`);
                }
                continue;
              }

              let date = null;
              // 삼성카드 형식: "20251101" (YYYYMMDD)
              if (dateStr.match(/^\d{8}$/)) {
                const year = dateStr.substring(0, 4);
                const month = dateStr.substring(4, 6);
                const day = dateStr.substring(6, 8);
                date = `${year}-${month}-${day}`;
              } else if (
                dateStr.includes("년") &&
                dateStr.includes("월") &&
                dateStr.includes("일")
              ) {
                date = parseKoreanDate(dateStr);
              } else if (
                !isNaN(parseFloat(dateStr)) &&
                parseFloat(dateStr) > 25569
              ) {
                // 엑셀 날짜 숫자 형식
                try {
                  const excelDateNum = parseFloat(dateStr);
                  const excelDate = XLSX.SSF.parse_date_code(excelDateNum);
                  if (excelDate) {
                    const year = excelDate.y;
                    const month = String(excelDate.m).padStart(2, "0");
                    const day = String(excelDate.d).padStart(2, "0");
                    date = `${year}-${month}-${day}`;
                  }
                } catch (e) {
                  // 파싱 실패 시 다음 방법 시도
                }
              }

              if (!date) {
                // 일반 날짜 문자열 처리
                const dateObj = new Date(dateStr);
                if (!isNaN(dateObj.getTime())) {
                  date = dateObj.toISOString().split("T")[0];
                }
              }

              if (!date) {
                if (i < headerRowIndex + 5) {
                  console.log(`행 ${i}: 날짜 파싱 실패 - "${dateStr}"`);
                }
                continue;
              }

              const category = categorizeByMerchant(merchantStr);

              transactions.push({
                id: Date.now() + i,
                type: "expense",
                category: category,
                amount: principalAmount,
                description: merchantStr,
                date: date,
                account:
                  selectedAccount ||
                  (config.accounts.length > 0 ? config.accounts[0] : ""),
              });

              continue;
            }
            // 국민은행 형식 처리
            else if (isKbBankFormat) {
              if (
                dateIndex === -1 ||
                (merchantIndex === -1 &&
                  summaryIndex === -1 &&
                  senderReceiverIndex === -1) ||
                (withdrawalIndex === -1 && depositIndex === -1)
              ) {
                if (i < headerRowIndex + 5) {
                  console.log(`행 ${i}: 필수 컬럼 없음`, {
                    dateIndex,
                    merchantIndex,
                    summaryIndex,
                    senderReceiverIndex,
                    withdrawalIndex,
                    depositIndex,
                  });
                }
                continue;
              }

              const dateStr = String(row[dateIndex] || "").trim();
              // 적요가 있으면 사용, 없으면 보낸분/받는분 사용
              let merchantStr = "";
              if (summaryIndex >= 0 && row[summaryIndex]) {
                merchantStr = String(row[summaryIndex]).trim();
              }
              if (
                !merchantStr &&
                senderReceiverIndex >= 0 &&
                row[senderReceiverIndex]
              ) {
                merchantStr = String(row[senderReceiverIndex]).trim();
              }
              if (!merchantStr && merchantIndex >= 0 && row[merchantIndex]) {
                merchantStr = String(row[merchantIndex]).trim();
              }
              const withdrawalStr =
                withdrawalIndex >= 0
                  ? String(row[withdrawalIndex] || "").trim()
                  : "";
              const depositStr =
                depositIndex >= 0 ? String(row[depositIndex] || "").trim() : "";

              if (i < headerRowIndex + 5) {
                console.log(`행 ${i} 데이터:`, {
                  dateStr,
                  merchantStr,
                  withdrawalStr,
                  depositStr,
                });
              }

              // 합계 행 건너뛰기 (어떤 컬럼에든 "합계"가 있으면 건너뛰기)
              const isSummaryRow = row.some((cell) => {
                const cellStr = String(cell || "").trim();
                return (
                  cellStr === "합계" ||
                  cellStr === "소계" ||
                  cellStr.includes("합계") ||
                  cellStr.includes("소계")
                );
              });

              // 소계, 합계, 빈 날짜 건너뛰기
              if (
                !dateStr ||
                dateStr === "-" ||
                isSummaryRow ||
                merchantStr.includes("소계") ||
                merchantStr.includes("합계") ||
                (merchantStr.includes("건") && !merchantStr.match(/\d+건/))
              ) {
                if (i < headerRowIndex + 5) {
                  console.log(`행 ${i}: 건너뛰기 (합계/소계 행 또는 빈 날짜)`);
                }
                continue;
              }

              // 출금액과 입금액 모두 없거나 0이면 건너뛰기
              const withdrawalNum = withdrawalStr
                ? parseFloat(withdrawalStr.toString().replace(/,/g, ""))
                : 0;
              const depositNum = depositStr
                ? parseFloat(depositStr.toString().replace(/,/g, ""))
                : 0;
              if (
                (!withdrawalStr || withdrawalNum === 0) &&
                (!depositStr || depositNum === 0)
              ) {
                if (i < headerRowIndex + 5) {
                  console.log(`행 ${i}: 건너뛰기 (출금액과 입금액 모두 없음)`);
                }
                continue;
              }

              let date = null;
              // 국민은행 형식: "2025.12.22 15:42:08" 또는 "2025.12.22"
              const dateTimeMatch = dateStr.match(
                /(\d{4})\.(\d{1,2})\.(\d{1,2})/
              );
              if (dateTimeMatch) {
                const year = dateTimeMatch[1];
                const month = dateTimeMatch[2].padStart(2, "0");
                const day = dateTimeMatch[3].padStart(2, "0");
                date = `${year}-${month}-${day}`;
              } else if (
                dateStr.includes("년") &&
                dateStr.includes("월") &&
                dateStr.includes("일")
              ) {
                date = parseKoreanDate(dateStr);
              } else if (
                !isNaN(parseFloat(dateStr)) &&
                parseFloat(dateStr) > 25569
              ) {
                // 엑셀 날짜 숫자 형식
                try {
                  const excelDateNum = parseFloat(dateStr);
                  const excelDate = XLSX.SSF.parse_date_code(excelDateNum);
                  if (excelDate) {
                    const year = excelDate.y;
                    const month = String(excelDate.m).padStart(2, "0");
                    const day = String(excelDate.d).padStart(2, "0");
                    date = `${year}-${month}-${day}`;
                  }
                } catch (e) {
                  // 파싱 실패 시 다음 방법 시도
                }
              }

              if (!date) {
                // 일반 날짜 문자열 처리
                const dateObj = new Date(dateStr);
                if (!isNaN(dateObj.getTime())) {
                  date = dateObj.toISOString().split("T")[0];
                }
              }

              if (!date) {
                if (i < headerRowIndex + 5) {
                  console.log(`행 ${i}: 날짜 파싱 실패 - "${dateStr}"`);
                }
                continue;
              }

              // 출금액과 입금액 처리 (이미 위에서 확인했으므로 여기서는 파싱만)
              const withdrawalAmount = withdrawalNum > 0 ? withdrawalNum : 0;
              const depositAmount = depositNum > 0 ? depositNum : 0;

              if (i < headerRowIndex + 5) {
                console.log(`행 ${i} 최종 금액:`, {
                  withdrawalAmount,
                  depositAmount,
                  date,
                  merchantStr,
                });
              }

              // 출금액이 있으면 지출, 입금액이 있으면 수입
              if (withdrawalAmount > 0) {
                const category = categorizeByMerchant(merchantStr);
                transactions.push({
                  id: Date.now() + i,
                  type: "expense",
                  category: category,
                  amount: withdrawalAmount,
                  description: merchantStr,
                  date: date,
                  account:
                    selectedAccount ||
                    (config.accounts.length > 0 ? config.accounts[0] : ""),
                });
              }

              if (depositAmount > 0) {
                // 입금은 수입으로 처리, 카테고리는 적요 내용에 따라 자동 분류하거나 기본값 사용
                let category = "기타 수입";
                const merchantLower = merchantStr.toLowerCase();
                if (
                  merchantLower.includes("급여") ||
                  merchantLower.includes("월급")
                ) {
                  category = "급여";
                } else if (merchantLower.includes("용돈")) {
                  category = "용돈";
                } else if (
                  merchantLower.includes("이자") ||
                  merchantLower.includes("예금")
                ) {
                  category = "부수입";
                }

                transactions.push({
                  id: Date.now() + i + 1000000, // 출금과 구분하기 위해 큰 수 추가
                  type: "income",
                  category: category,
                  amount: depositAmount,
                  description: merchantStr,
                  date: date,
                  account:
                    selectedAccount ||
                    (config.accounts.length > 0 ? config.accounts[0] : ""),
                });
              }

              continue;
            }

            // 우리은행 형식 처리
            if (isWooriBankFormat) {
              if (
                dateIndex === -1 ||
                (merchantIndex === -1 && summaryIndex === -1) ||
                (withdrawalIndex === -1 && depositIndex === -1)
              ) {
                continue;
              }

              const dateStr = String(row[dateIndex] || "").trim();
              // 기재내용이 있으면 사용, 없으면 적요 사용
              let merchantStr = "";
              if (merchantIndex >= 0 && row[merchantIndex]) {
                merchantStr = String(row[merchantIndex]).trim();
              }
              if (!merchantStr && summaryIndex >= 0 && row[summaryIndex]) {
                merchantStr = String(row[summaryIndex]).trim();
              }
              const withdrawalStr =
                withdrawalIndex >= 0
                  ? String(row[withdrawalIndex] || "").trim()
                  : "";
              const depositStr =
                depositIndex >= 0 ? String(row[depositIndex] || "").trim() : "";

              // 소계, 합계 건너뛰기
              if (
                !dateStr ||
                merchantStr.includes("소계") ||
                merchantStr.includes("합계") ||
                (merchantStr.includes("건") && !merchantStr.match(/\d+건/))
              ) {
                continue;
              }

              // 출금액과 입금액 모두 없으면 건너뛰기
              if (!withdrawalStr && !depositStr) continue;

              let date = null;
              // 날짜 형식 변환 시도
              // 우리은행 형식: "2025.11.30 13:20" 또는 "2025.11.30"
              const dateTimeMatch = dateStr.match(
                /(\d{4})\.(\d{1,2})\.(\d{1,2})/
              );
              if (dateTimeMatch) {
                const year = dateTimeMatch[1];
                const month = dateTimeMatch[2].padStart(2, "0");
                const day = dateTimeMatch[3].padStart(2, "0");
                date = `${year}-${month}-${day}`;
              } else if (
                dateStr.includes("년") &&
                dateStr.includes("월") &&
                dateStr.includes("일")
              ) {
                date = parseKoreanDate(dateStr);
              } else if (
                !isNaN(parseFloat(dateStr)) &&
                parseFloat(dateStr) > 25569
              ) {
                // 엑셀 날짜 숫자 형식
                try {
                  const excelDateNum = parseFloat(dateStr);
                  const excelDate = XLSX.SSF.parse_date_code(excelDateNum);
                  if (excelDate) {
                    const year = excelDate.y;
                    const month = String(excelDate.m).padStart(2, "0");
                    const day = String(excelDate.d).padStart(2, "0");
                    date = `${year}-${month}-${day}`;
                  }
                } catch (e) {
                  // 파싱 실패 시 다음 방법 시도
                }
              }

              if (!date) {
                // 일반 날짜 문자열 처리
                const dateObj = new Date(dateStr);
                if (!isNaN(dateObj.getTime())) {
                  date = dateObj.toISOString().split("T")[0];
                }
              }

              if (!date) continue;

              // 출금액과 입금액 처리 (0이나 빈 문자열도 처리)
              const withdrawalAmount =
                withdrawalStr &&
                withdrawalStr !== "0" &&
                withdrawalStr.trim() !== ""
                  ? parseFloat(withdrawalStr.toString().replace(/,/g, ""))
                  : 0;
              const depositAmount =
                depositStr && depositStr !== "0" && depositStr.trim() !== ""
                  ? parseFloat(depositStr.toString().replace(/,/g, ""))
                  : 0;

              // 출금액이 있으면 지출, 입금액이 있으면 수입
              if (withdrawalAmount > 0) {
                const category = categorizeByMerchant(merchantStr);
                transactions.push({
                  id: Date.now() + i,
                  type: "expense",
                  category: category,
                  amount: withdrawalAmount,
                  description: merchantStr,
                  date: date,
                  account:
                    selectedAccount ||
                    (config.accounts.length > 0 ? config.accounts[0] : ""),
                });
              }

              if (depositAmount > 0) {
                // 입금은 수입으로 처리, 카테고리는 적요 내용에 따라 자동 분류하거나 기본값 사용
                let category = "기타 수입";
                const merchantLower = merchantStr.toLowerCase();
                if (
                  merchantLower.includes("급여") ||
                  merchantLower.includes("월급")
                ) {
                  category = "급여";
                } else if (merchantLower.includes("용돈")) {
                  category = "용돈";
                } else if (
                  merchantLower.includes("이자") ||
                  merchantLower.includes("예금")
                ) {
                  category = "부수입";
                }

                transactions.push({
                  id: Date.now() + i + 1000000, // 출금과 구분하기 위해 큰 수 추가
                  type: "income",
                  category: category,
                  amount: depositAmount,
                  description: merchantStr,
                  date: date,
                  account:
                    selectedAccount ||
                    (config.accounts.length > 0 ? config.accounts[0] : ""),
                });
              }

              continue;
            }

            // 기존 형식 처리 (현대카드 등)
            if (!row[dateIndex] || !row[merchantIndex] || !row[amountIndex])
              continue;

            const dateStr = String(row[dateIndex]).trim();
            const merchantStr = String(row[merchantIndex]).trim();
            const amountStr = String(row[amountIndex]).trim();

            // 소계, 합계 건너뛰기 (단, "0003건" 같은 패턴은 제외)
            if (
              merchantStr.includes("소계") ||
              merchantStr.includes("합계") ||
              (merchantStr.includes("건") && !merchantStr.match(/\d+건/)) ||
              amountStr === ""
            ) {
              continue;
            }

            let date = null;
            // 날짜 형식 변환 시도
            if (
              dateStr.includes("년") &&
              dateStr.includes("월") &&
              dateStr.includes("일")
            ) {
              date = parseKoreanDate(dateStr);
            } else if (
              !isNaN(parseFloat(dateStr)) &&
              parseFloat(dateStr) > 25569
            ) {
              // 엑셀 날짜 숫자 형식 (1900-01-01 이후)
              try {
                const excelDateNum = parseFloat(dateStr);
                const excelDate = XLSX.SSF.parse_date_code(excelDateNum);
                if (excelDate) {
                  const year = excelDate.y;
                  const month = String(excelDate.m).padStart(2, "0");
                  const day = String(excelDate.d).padStart(2, "0");
                  date = `${year}-${month}-${day}`;
                }
              } catch (e) {
                // 파싱 실패 시 다음 방법 시도
              }
            }

            if (!date) {
              // 일반 날짜 문자열 처리
              const dateObj = new Date(dateStr);
              if (!isNaN(dateObj.getTime())) {
                date = dateObj.toISOString().split("T")[0];
              }
            }

            if (!date) continue;

            let amount = parseFloat(amountStr.toString().replace(/,/g, ""));
            // 금액이 0이거나 유효하지 않으면 건너뛰기
            if (isNaN(amount) || amount <= 0) continue;

            const category = categorizeByMerchant(merchantStr);

            transactions.push({
              id: Date.now() + i,
              type: "expense",
              category: category,
              amount: amount,
              description: merchantStr,
              date: date,
              account:
                selectedAccount ||
                (config.accounts.length > 0 ? config.accounts[0] : ""),
            });
          }

          console.log("국민은행 파싱 완료:", transactions.length, "개 내역");
          resolve(transactions);
        } catch (error) {
          reject(error);
        }
      };

      reader.onerror = (error) => {
        console.error("파일 읽기 오류:", error);
        reject(
          new Error(
            "파일을 읽을 수 없습니다. 파일이 손상되었거나 지원하지 않는 형식일 수 있습니다."
          )
        );
      };

      reader.onabort = () => {
        console.error("파일 읽기 중단됨");
        reject(new Error("파일 읽기가 중단되었습니다."));
      };

      try {
        reader.readAsArrayBuffer(file);
      } catch (error) {
        console.error("파일 읽기 시작 오류:", error);
        reject(new Error("파일을 읽을 수 없습니다: " + error.message));
      }
    });
  };

  // 계좌 선택 모달 열기
  const handleOpenAccountModal = () => {
    setShowAccountModal(true);
    setSelectedAccountForUpload(
      config.accounts.length > 0 ? config.accounts[0] : ""
    );
  };

  // 계좌 선택 후 파일 선택 진행
  const handleAccountSelected = () => {
    if (!selectedAccountForUpload) {
      alert("계좌를 선택해주세요.");
      return;
    }
    setShowAccountModal(false);
    // 파일 입력 트리거
    if (fileInputRef.current) {
      fileInputRef.current.click();
    }
  };

  // 파일 업로드 처리
  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const selectedAccount =
      selectedAccountForUpload ||
      formData.account ||
      (config.accounts.length > 0 ? config.accounts[0] : "");

    const fileName = file.name.toLowerCase();
    let transactions = [];

    try {
      // 파일이 HTML인지 엑셀인지 확인
      const checkFileType = (file) => {
        return new Promise((resolve) => {
          if (
            fileName.endsWith(".html") ||
            fileName.endsWith(".htm") ||
            file.type === "text/html"
          ) {
            resolve("html");
            return;
          }

          // .xls 파일인 경우 내용을 확인
          if (fileName.endsWith(".xls") && !fileName.endsWith(".xlsx")) {
            const reader = new FileReader();
            reader.onload = (e) => {
              const content = e.target.result;
              // HTML 태그가 있으면 HTML 파일
              if (
                typeof content === "string" &&
                (content.includes("<html") ||
                  content.includes("<table") ||
                  content.includes("<!DOCTYPE"))
              ) {
                resolve("html");
              } else {
                resolve("excel");
              }
            };
            reader.onerror = () => resolve("excel");
            reader.readAsText(file.slice(0, 1024), "utf-8"); // 첫 1KB만 읽어서 확인
          } else if (fileName.endsWith(".xlsx")) {
            resolve("excel");
          } else {
            resolve("html"); // 기본값
          }
        });
      };

      const fileType = await checkFileType(file);
      console.log(
        "파일 타입 감지:",
        fileType,
        "파일명:",
        fileName,
        "MIME 타입:",
        file.type
      );

      if (fileType === "html") {
        // HTML 파일 처리
        const reader = new FileReader();
        reader.onload = (e) => {
          try {
            const htmlContent = e.target.result;
            console.log("HTML 파일 읽기 완료, 파싱 시작...");
            transactions = parseHTMLFile(htmlContent, selectedAccount);

            if (transactions.length > 0) {
              const confirmMsg = `${transactions.length}개의 내역을 추가하시겠습니까?`;
              if (window.confirm(confirmMsg)) {
                const newTransactions = transactions.map((t) => ({
                  ...t,
                  id: Date.now() + Math.random(),
                }));
                setTransactions((prev) => [...newTransactions, ...prev]);
                alert(`${transactions.length}개의 내역이 추가되었습니다.`);
              }
            } else {
              console.error("파싱된 내역이 없습니다.");
              alert(
                "파일에서 내역을 찾을 수 없습니다.\n\n브라우저 개발자 도구(F12)의 콘솔 탭에서 상세한 오류 정보를 확인할 수 있습니다.\n\n파일 형식을 확인해주세요.\n(현대카드 명세서 HTML 형식, 삼성카드/국민은행/우리은행 거래내역 엑셀 파일 또는 일반 엑셀 파일)"
              );
            }
          } catch (error) {
            console.error("파일 처리 오류:", error);
            alert("파일을 처리하는 중 오류가 발생했습니다: " + error.message);
          }
        };
        reader.onerror = () => {
          alert("파일을 읽는 중 오류가 발생했습니다.");
        };
        reader.readAsText(file, "utf-8");
      } else if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
        // 엑셀 파일 처리
        console.log("엑셀 파일 처리 시작...");
        transactions = await parseExcelFile(file, selectedAccount);
        if (transactions.length > 0) {
          const confirmMsg = `${transactions.length}개의 내역을 추가하시겠습니까?`;
          if (window.confirm(confirmMsg)) {
            const newTransactions = transactions.map((t) => ({
              ...t,
              id: Date.now() + Math.random(),
            }));
            setTransactions((prev) => [...newTransactions, ...prev]);
            alert(`${transactions.length}개의 내역이 추가되었습니다.`);
          }
        } else {
          console.error("파싱된 내역이 없습니다.");
          alert(
            "파일에서 내역을 찾을 수 없습니다.\n\n브라우저 개발자 도구(F12)의 콘솔 탭에서 상세한 오류 정보를 확인할 수 있습니다.\n\n파일 형식을 확인해주세요.\n(국민은행, 우리은행 거래내역 엑셀 파일 또는 일반 엑셀 형식)"
          );
        }
      } else {
        alert(
          "지원하지 않는 파일 형식입니다. (.xls, .xlsx, .html 파일만 지원)"
        );
      }
    } catch (error) {
      console.error("파일 파싱 오류:", error);
      const errorMessage = error.message || "알 수 없는 오류가 발생했습니다.";
      alert(
        "파일을 읽는 중 오류가 발생했습니다:\n\n" +
          errorMessage +
          "\n\n브라우저 개발자 도구(F12)의 콘솔 탭에서 상세한 오류 정보를 확인할 수 있습니다."
      );
    }

    // 파일 입력 초기화 및 선택된 계좌 초기화
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
    setSelectedAccountForUpload("");
  };

  const filteredTransactions = transactions
    .filter((t) => {
      if (filter === "income" && t.type !== "income") return false;
      if (filter === "expense" && t.type !== "expense") return false;
      if (startDate && t.date < startDate) return false;
      if (endDate && t.date > endDate) return false;
      return true;
    })
    .sort((a, b) => {
      // 정렬 기준에 따라 정렬
      let compareA, compareB;

      if (sortBy === "date") {
        compareA = a.date;
        compareB = b.date;
      } else if (sortBy === "account") {
        compareA = a.account || "";
        compareB = b.account || "";
      } else if (sortBy === "category") {
        compareA = a.category || "";
        compareB = b.category || "";
      } else {
        return 0;
      }

      // 정렬 순서에 따라 비교
      if (sortOrder === "asc") {
        if (compareA < compareB) return -1;
        if (compareA > compareB) return 1;
        return 0;
      } else {
        if (compareA > compareB) return -1;
        if (compareA < compareB) return 1;
        return 0;
      }
    });

  const totalIncome = transactions
    .filter((t) => t.type === "income")
    .reduce((sum, t) => sum + t.amount, 0);

  const totalExpense = transactions
    .filter((t) => t.type === "expense")
    .reduce((sum, t) => sum + t.amount, 0);

  const balance = totalIncome - totalExpense;

  // 엑셀 파일로 내보내기
  const handleExportToExcel = () => {
    // 기본 파일명 설정
    const defaultFileName = `가계부_${
      new Date().toISOString().split("T")[0]
    }.xlsx`;
    setExportFileName(defaultFileName.replace(".xlsx", ""));
    setShowFileNameModal(true);
  };

  // 실제 엑셀 내보내기 실행
  const executeExportToExcel = () => {
    try {
      // 워크북 생성
      const workbook = XLSX.utils.book_new();

      // 계좌별 수입/지출/잔액 계산
      const accountSummary = {};
      filteredTransactions.forEach((transaction) => {
        const account = transaction.account || "미지정";
        if (!accountSummary[account]) {
          accountSummary[account] = { income: 0, expense: 0 };
        }
        if (transaction.type === "income") {
          accountSummary[account].income += transaction.amount;
        } else {
          accountSummary[account].expense += transaction.amount;
        }
      });

      // 계좌별 잔액 계산
      const accountBalances = Object.keys(accountSummary).map((account) => ({
        account,
        income: accountSummary[account].income,
        expense: accountSummary[account].expense,
        balance:
          accountSummary[account].income - accountSummary[account].expense,
      }));

      // 요약 정보 시트
      const summaryData = [
        ["가계부 내역 내보내기"],
        [""],
        ["내보낸 날짜", new Date().toLocaleString("ko-KR")],
        [""],
        ["전체 요약 정보"],
        ["수입", totalIncome.toLocaleString() + "원"],
        ["지출", totalExpense.toLocaleString() + "원"],
        ["잔액", balance.toLocaleString() + "원"],
        [""],
        ["계좌별 요약 정보"],
        ["계좌", "수입", "지출", "잔액"],
      ];

      // 계좌별 정보 추가
      accountBalances.forEach(({ account, income, expense, balance }) => {
        summaryData.push([
          account,
          income.toLocaleString() + "원",
          expense.toLocaleString() + "원",
          balance.toLocaleString() + "원",
        ]);
      });

      summaryData.push([""]);
      summaryData.push(["필터 정보"]);
      summaryData.push([
        "유형",
        filter === "all" ? "전체" : filter === "income" ? "수입" : "지출",
      ]);
      summaryData.push(["시작 날짜", startDate || "전체"]);
      summaryData.push(["종료 날짜", endDate || "전체"]);
      summaryData.push([
        "정렬 기준",
        sortBy === "date"
          ? "날짜순"
          : sortBy === "account"
          ? "계좌순"
          : "카테고리순",
      ]);
      summaryData.push([
        "정렬 순서",
        sortOrder === "asc" ? "오름차순" : "내림차순",
      ]);
      summaryData.push([""]);
      summaryData.push(["총 내역 수", filteredTransactions.length + "건"]);

      const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);

      // 요약 시트 컬럼 너비 설정
      summarySheet["!cols"] = [
        { wch: 20 }, // 첫 번째 컬럼
        { wch: 20 }, // 두 번째 컬럼
        { wch: 20 }, // 세 번째 컬럼 (계좌별 정보)
        { wch: 20 }, // 네 번째 컬럼 (계좌별 정보)
      ];

      XLSX.utils.book_append_sheet(workbook, summarySheet, "요약");

      // 거래 내역 시트
      const transactionData = [
        ["날짜", "유형", "카테고리", "서브카테고리", "계좌", "금액", "내용"],
      ];

      filteredTransactions.forEach((transaction) => {
        transactionData.push([
          transaction.date,
          transaction.type === "income" ? "수입" : "지출",
          transaction.category || "",
          transaction.subCategory || "",
          transaction.account || "",
          transaction.amount,
          transaction.description || "",
        ]);
      });

      const transactionSheet = XLSX.utils.aoa_to_sheet(transactionData);

      // 컬럼 너비 설정
      const colWidths = [
        { wch: 12 }, // 날짜
        { wch: 8 }, // 유형
        { wch: 12 }, // 카테고리
        { wch: 15 }, // 서브카테고리
        { wch: 12 }, // 계좌
        { wch: 15 }, // 금액
        { wch: 30 }, // 내용
      ];
      transactionSheet["!cols"] = colWidths;

      XLSX.utils.book_append_sheet(workbook, transactionSheet, "거래 내역");

      // 파일명 생성 (사용자 입력 또는 기본값)
      let fileName = exportFileName.trim();
      if (!fileName) {
        fileName = `가계부_${new Date().toISOString().split("T")[0]}`;
      }
      // .xlsx 확장자가 없으면 추가
      if (!fileName.endsWith(".xlsx") && !fileName.endsWith(".xls")) {
        fileName += ".xlsx";
      }

      // 엑셀 파일 다운로드
      XLSX.writeFile(workbook, fileName);

      alert(`엑셀 파일이 다운로드되었습니다.\n파일명: ${fileName}`);
      setShowFileNameModal(false);
      setExportFileName("");
    } catch (error) {
      console.error("엑셀 내보내기 오류:", error);
      alert("엑셀 파일 내보내기 중 오류가 발생했습니다: " + error.message);
      setShowFileNameModal(false);
      setExportFileName("");
    }
  };

  return (
    <div className="app">
      <header className="header">
        <h1>💰 가계부</h1>
        <div className="summary">
          <div className="summary-item income">
            <span className="label">수입</span>
            <span className="amount">+{totalIncome.toLocaleString()}원</span>
          </div>
          <div className="summary-item expense">
            <span className="label">지출</span>
            <span className="amount">-{totalExpense.toLocaleString()}원</span>
          </div>
          <div className="summary-item balance">
            <span className="label">잔액</span>
            <span
              className={`amount ${balance >= 0 ? "positive" : "negative"}`}
            >
              {balance.toLocaleString()}원
            </span>
          </div>
        </div>
        <div className="file-management-summary">
          <div className="file-management-item">
            <div className="file-management-icon">➕</div>
            <div className="file-management-content">
              <span className="file-management-title">거래 내역 추가</span>
              <span className="file-management-desc">
                카드사 거래내역 파일 업로드
              </span>
            </div>
            <button
              type="button"
              onClick={handleOpenAccountModal}
              className="file-management-button"
            >
              추가하기
            </button>
            <input
              id="file-upload"
              ref={fileInputRef}
              type="file"
              accept=".xls,.xlsx,.html,.htm"
              onChange={handleFileUpload}
              className="file-upload-input"
            />
          </div>
          <div className="file-management-item">
            <div className="file-management-icon">🔄</div>
            <div className="file-management-content">
              <span className="file-management-title">파일 교체</span>
              <span className="file-management-desc warning">
                ⚠️ 기존 데이터 삭제됨
              </span>
            </div>
            <label
              htmlFor="main-file-reload"
              className="file-management-button"
            >
              파일 선택
            </label>
            <input
              id="main-file-reload"
              ref={mainFileInputRef}
              type="file"
              accept=".xlsx,.xls"
              onChange={handleLoadMainFile}
              className="file-upload-input"
            />
          </div>
        </div>
      </header>

      <main className="main">
        <section className="form-section">
          <h2>내역 추가</h2>

          <form onSubmit={handleSubmit} className="transaction-form">
            <div className="form-group">
              <label>유형</label>
              <div className="type-buttons">
                <button
                  type="button"
                  className={formData.type === "income" ? "active income" : ""}
                  onClick={() =>
                    setFormData({
                      ...formData,
                      type: "income",
                      category: "",
                    })
                  }
                >
                  수입
                </button>
                <button
                  type="button"
                  className={
                    formData.type === "expense" ? "active expense" : ""
                  }
                  onClick={() =>
                    setFormData({
                      ...formData,
                      type: "expense",
                      category: "",
                    })
                  }
                >
                  지출
                </button>
              </div>
            </div>

            <div className="form-group">
              <label>카테고리</label>
              <select
                value={formData.category}
                onChange={(e) =>
                  setFormData({
                    ...formData,
                    category: e.target.value,
                    subCategory: "",
                  })
                }
                required
              >
                <option value="">선택하세요</option>
                {categories[formData.type].map((cat) => (
                  <option key={cat} value={cat}>
                    {cat}
                  </option>
                ))}
              </select>
            </div>

            {formData.category &&
              config.subCategories[formData.category] &&
              config.subCategories[formData.category].length > 0 && (
                <div className="form-group">
                  <label>서브 카테고리</label>
                  <select
                    value={formData.subCategory}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        subCategory: e.target.value,
                      })
                    }
                  >
                    <option value="">선택 안함</option>
                    {config.subCategories[formData.category].map((subCat) => (
                      <option key={subCat} value={subCat}>
                        {subCat}
                      </option>
                    ))}
                  </select>
                </div>
              )}

            <div className="form-group">
              <label>금액</label>
              <input
                type="number"
                value={formData.amount}
                onChange={(e) =>
                  setFormData({ ...formData, amount: e.target.value })
                }
                placeholder="금액을 입력하세요"
                min="0"
                step="100"
                required
              />
            </div>

            <div className="form-group">
              <label>내용</label>
              <input
                type="text"
                value={formData.description}
                onChange={(e) =>
                  setFormData({ ...formData, description: e.target.value })
                }
                placeholder="내용을 입력하세요"
                required
              />
            </div>

            <div className="form-group">
              <label>계좌</label>
              <select
                value={formData.account}
                onChange={(e) =>
                  setFormData({ ...formData, account: e.target.value })
                }
                required
              >
                <option value="">선택하세요</option>
                {config.accounts.map((account) => (
                  <option key={account} value={account}>
                    {account}
                  </option>
                ))}
              </select>
            </div>

            <div className="form-group">
              <label>날짜</label>
              <input
                type="date"
                value={formData.date}
                onChange={(e) =>
                  setFormData({ ...formData, date: e.target.value })
                }
                required
              />
            </div>

            <button type="submit" className="submit-btn">
              추가하기
            </button>
          </form>
        </section>

        <section className="list-section">
          <div className="filter-controls">
            <div className="list-header">
              <h2>내역 목록</h2>
              <button
                onClick={handleDeleteAll}
                className="delete-all-btn"
                title="모든 거래 내역 삭제"
              >
                🗑️ 전체 삭제
              </button>
            </div>
            <div className="filters">
              <select
                value={filter}
                onChange={(e) => setFilter(e.target.value)}
                className="filter-select"
              >
                <option value="all">전체</option>
                <option value="income">수입</option>
                <option value="expense">지출</option>
              </select>
              <div className="sort-controls">
                <select
                  value={sortBy}
                  onChange={(e) => setSortBy(e.target.value)}
                  className="sort-select"
                >
                  <option value="date">날짜순</option>
                  <option value="account">계좌순</option>
                  <option value="category">카테고리순</option>
                </select>
                <button
                  type="button"
                  onClick={() =>
                    setSortOrder(sortOrder === "asc" ? "desc" : "asc")
                  }
                  className="sort-order-btn"
                  title={sortOrder === "asc" ? "오름차순" : "내림차순"}
                >
                  {sortOrder === "asc" ? "↑" : "↓"}
                </button>
              </div>
              <div className="date-range-filters">
                <input
                  type="date"
                  value={startDate}
                  onChange={(e) => setStartDate(e.target.value)}
                  className="date-filter"
                  placeholder="시작 날짜"
                />
                <span className="date-separator">~</span>
                <input
                  type="date"
                  value={endDate}
                  onChange={(e) => setEndDate(e.target.value)}
                  className="date-filter"
                  placeholder="종료 날짜"
                />
                {(startDate || endDate) && (
                  <button
                    onClick={() => {
                      setStartDate("");
                      setEndDate("");
                    }}
                    className="clear-filter"
                  >
                    필터 초기화
                  </button>
                )}
              </div>
              <button
                onClick={handleExportToExcel}
                className="export-excel-btn"
                title="엑셀 파일로 내보내기"
              >
                📊 엑셀 내보내기
              </button>
            </div>
          </div>

          <div className="transaction-list">
            {filteredTransactions.length === 0 ? (
              <div className="empty-state">
                <p>내역이 없습니다.</p>
              </div>
            ) : (
              filteredTransactions.map((transaction) => (
                <div
                  key={transaction.id}
                  className={`transaction-item ${transaction.type}`}
                >
                  <div className="transaction-info">
                    {editingTransaction === transaction.id ? (
                      <div className="transaction-edit-form">
                        <div className="edit-form-row">
                          <div className="edit-form-group">
                            <label>날짜</label>
                            <input
                              type="date"
                              value={editingTransactionData.date}
                              onChange={(e) =>
                                setEditingTransactionData({
                                  ...editingTransactionData,
                                  date: e.target.value,
                                })
                              }
                              className="transaction-edit-input"
                              required
                            />
                          </div>
                          <div className="edit-form-group">
                            <label>유형</label>
                            <select
                              value={editingTransactionData.type}
                              onChange={(e) =>
                                setEditingTransactionData({
                                  ...editingTransactionData,
                                  type: e.target.value,
                                  category: "",
                                  subCategory: "",
                                })
                              }
                              className="transaction-edit-select"
                              required
                            >
                              <option value="income">수입</option>
                              <option value="expense">지출</option>
                            </select>
                          </div>
                        </div>
                        <div className="edit-form-row">
                          <div className="edit-form-group">
                            <label>카테고리</label>
                            <select
                              value={editingTransactionData.category}
                              onChange={(e) =>
                                setEditingTransactionData({
                                  ...editingTransactionData,
                                  category: e.target.value,
                                  subCategory: "",
                                })
                              }
                              className="transaction-edit-select"
                              required
                            >
                              <option value="">선택하세요</option>
                              {config.categories[
                                editingTransactionData.type
                              ]?.map((cat) => (
                                <option key={cat} value={cat}>
                                  {cat}
                                </option>
                              ))}
                            </select>
                          </div>
                          <div className="edit-form-group">
                            <label>서브 카테고리</label>
                            <select
                              value={editingTransactionData.subCategory}
                              onChange={(e) =>
                                setEditingTransactionData({
                                  ...editingTransactionData,
                                  subCategory: e.target.value,
                                })
                              }
                              className="transaction-edit-select"
                            >
                              <option value="">선택 안함</option>
                              {config.subCategories[
                                editingTransactionData.category
                              ]?.map((subCat) => (
                                <option key={subCat} value={subCat}>
                                  {subCat}
                                </option>
                              ))}
                            </select>
                          </div>
                        </div>
                        <div className="edit-form-row">
                          <div className="edit-form-group">
                            <label>계좌</label>
                            <select
                              value={editingTransactionData.account}
                              onChange={(e) =>
                                setEditingTransactionData({
                                  ...editingTransactionData,
                                  account: e.target.value,
                                })
                              }
                              className="transaction-edit-select"
                              required
                            >
                              <option value="">선택하세요</option>
                              {config.accounts.map((account) => (
                                <option key={account} value={account}>
                                  {account}
                                </option>
                              ))}
                            </select>
                          </div>
                          <div className="edit-form-group">
                            <label>금액</label>
                            <input
                              type="number"
                              value={editingTransactionData.amount}
                              onChange={(e) =>
                                setEditingTransactionData({
                                  ...editingTransactionData,
                                  amount: e.target.value,
                                })
                              }
                              className="transaction-edit-input"
                              min="0"
                              step="100"
                              required
                            />
                          </div>
                        </div>
                        <div className="edit-form-row">
                          <div className="edit-form-group full-width">
                            <label>내용</label>
                            <input
                              type="text"
                              value={editingTransactionData.description}
                              onChange={(e) =>
                                setEditingTransactionData({
                                  ...editingTransactionData,
                                  description: e.target.value,
                                })
                              }
                              className="transaction-edit-input"
                              required
                            />
                          </div>
                        </div>
                      </div>
                    ) : (
                      <>
                        <div className="transaction-header">
                          <div className="category-group">
                            <span className="category">
                              {transaction.category}
                            </span>
                            {transaction.subCategory && (
                              <span className="sub-category">
                                {transaction.subCategory}
                              </span>
                            )}
                          </div>
                          <span className={`amount ${transaction.type}`}>
                            {transaction.type === "income" ? "+" : "-"}
                            {transaction.amount.toLocaleString()}원
                          </span>
                        </div>
                        <div className="transaction-details">
                          <span className="description">
                            {transaction.description}
                          </span>
                          <div className="transaction-meta">
                            {transaction.account && (
                              <span className="account">
                                {transaction.account}
                              </span>
                            )}
                            <span className="date">{transaction.date}</span>
                          </div>
                        </div>
                      </>
                    )}
                  </div>
                  <div className="transaction-actions">
                    {editingTransaction === transaction.id ? (
                      <>
                        <button
                          onClick={() => handleSaveTransaction(transaction.id)}
                          className="save-btn"
                          title="저장"
                        >
                          ✓
                        </button>
                        <button
                          onClick={handleCancelEdit}
                          className="cancel-btn"
                          title="취소"
                        >
                          ✕
                        </button>
                      </>
                    ) : (
                      <>
                        <button
                          onClick={() => handleEditTransaction(transaction)}
                          className="edit-btn"
                          title="수정"
                        >
                          ✏️
                        </button>
                        <button
                          onClick={() => handleDelete(transaction.id)}
                          className="delete-btn"
                          title="삭제"
                        >
                          🗑️
                        </button>
                      </>
                    )}
                  </div>
                </div>
              ))
            )}
          </div>
        </section>
      </main>

      {/* 파일명 입력 모달 */}
      {showFileNameModal && (
        <div
          className="modal-overlay"
          onClick={() => {
            setShowFileNameModal(false);
            setExportFileName("");
          }}
        >
          <div className="modal-content" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3>파일명 입력</h3>
              <button
                className="modal-close"
                onClick={() => {
                  setShowFileNameModal(false);
                  setExportFileName("");
                }}
              >
                ×
              </button>
            </div>
            <div className="modal-body">
              <p className="modal-description">
                내보낼 엑셀 파일의 이름을 입력해주세요.
              </p>
              <input
                type="text"
                className="file-name-input"
                value={exportFileName}
                onChange={(e) => setExportFileName(e.target.value)}
                placeholder="파일명을 입력하세요 (확장자 제외)"
                autoFocus
                onKeyDown={(e) => {
                  if (e.key === "Enter") {
                    executeExportToExcel();
                  } else if (e.key === "Escape") {
                    setShowFileNameModal(false);
                    setExportFileName("");
                  }
                }}
              />
              <p className="file-name-hint">
                확장자(.xlsx)는 자동으로 추가됩니다.
              </p>
            </div>
            <div className="modal-footer">
              <button
                type="button"
                className="modal-button cancel"
                onClick={() => {
                  setShowFileNameModal(false);
                  setExportFileName("");
                }}
              >
                취소
              </button>
              <button
                type="button"
                className="modal-button confirm"
                onClick={executeExportToExcel}
              >
                내보내기
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 계좌 선택 모달 */}
      {showAccountModal && (
        <div
          className="modal-overlay"
          onClick={() => setShowAccountModal(false)}
        >
          <div className="modal-content" onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3>계좌 선택</h3>
              <button
                className="modal-close"
                onClick={() => setShowAccountModal(false)}
              >
                ×
              </button>
            </div>
            <div className="modal-body">
              <p className="modal-description">
                거래내역을 추가할 계좌를 선택해주세요.
              </p>
              <div className="account-select-list">
                {config.accounts.map((account) => (
                  <button
                    key={account}
                    type="button"
                    className={`account-select-item ${
                      selectedAccountForUpload === account ? "selected" : ""
                    }`}
                    onClick={() => setSelectedAccountForUpload(account)}
                  >
                    {account}
                  </button>
                ))}
              </div>
            </div>
            <div className="modal-footer">
              <button
                type="button"
                className="modal-button cancel"
                onClick={() => setShowAccountModal(false)}
              >
                취소
              </button>
              <button
                type="button"
                className="modal-button confirm"
                onClick={handleAccountSelected}
              >
                확인
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
