// 전역 변수
let studentsList = [];
let templateFile = null;

// DOM이 로드된 후 실행
document.addEventListener("DOMContentLoaded", function () {
  console.log("Script loaded");

  // DOM 요소
  const templateFileInput = document.getElementById("templateFile");
  const templateFileName = document.getElementById("templateFileName");
  const selectTemplateBtn = document.getElementById("selectTemplateBtn");
  const certiNumStart = document.getElementById("certiNumStart");
  const subName = document.getElementById("subName");
  const location = document.getElementById("location");
  const eduDate = document.getElementById("eduDate");
  const studentInput = document.getElementById("studentInput");
  const addStudentBtn = document.getElementById("addStudentBtn");
  const clearBtn = document.getElementById("clearBtn");
  const studentList = document.getElementById("studentList");
  const exportExcelBtn = document.getElementById("exportExcelBtn");
  const exportWordBtn = document.getElementById("exportWordBtn");

  // 요소 존재 확인
  if (
    !templateFileInput ||
    !selectTemplateBtn ||
    !addStudentBtn ||
    !clearBtn ||
    !exportExcelBtn ||
    !exportWordBtn ||
    !studentInput ||
    !studentList
  ) {
    console.error("필수 DOM 요소를 찾을 수 없습니다.");
    return;
  }

  // 이벤트 리스너
  selectTemplateBtn.addEventListener("click", () => templateFileInput.click());
  templateFileInput.addEventListener("change", handleTemplateFileSelect);
  addStudentBtn.addEventListener("click", addStudent);
  clearBtn.addEventListener("click", clearInput);
  exportExcelBtn.addEventListener("click", exportToExcel);
  exportWordBtn.addEventListener("click", exportToWord);

  // Enter 키로 학생 추가 (Ctrl+Enter)
  studentInput.addEventListener("keydown", (e) => {
    if (e.ctrlKey && e.key === "Enter") {
      addStudent();
    }
  });

  // 템플릿 파일 선택
  function handleTemplateFileSelect(event) {
    const file = event.target.files[0];
    if (file) {
      templateFile = file;
      templateFileName.value = file.name;
    }
  }

  // 입력창 초기화
  function clearInput() {
    studentInput.value = "";
    studentInput.focus();
  }

  // 학생 추가
  function addStudent() {
    const inputText = studentInput.value.trim();

    if (!inputText) {
      alert("학생명을 입력해주세요.");
      return;
    }

    // 여러 줄을 개행 문자로 분리
    const studentNames = inputText
      .split("\n")
      .map((name) => name.trim())
      .filter((name) => name.length > 0);

    if (studentNames.length === 0) {
      alert("학생명을 입력해주세요.");
      return;
    }

    let addedCount = 0;
    let skippedCount = 0;
    const errorMessages = [];

    studentNames.forEach((studentName) => {
      // 길이 검증
      if (studentName.length > 30) {
        errorMessages.push(`'${studentName}'는 30글자를 초과합니다.`);
        skippedCount++;
        return;
      }

      // 중복 검증
      if (studentsList.includes(studentName)) {
        skippedCount++;
        return;
      }

      // 리스트에 추가
      studentsList.push(studentName);
      addedCount++;
    });

    // 리스트 업데이트
    updateStudentList();
    clearInput();

    // 결과 메시지
    if (addedCount > 0) {
      if (skippedCount > 0 || errorMessages.length > 0) {
        let msg = `${addedCount}명의 학생이 추가되었습니다.`;
        if (skippedCount > 0) {
          msg += `\n${skippedCount}명은 중복 또는 오류로 제외되었습니다.`;
        }
        if (errorMessages.length > 0) {
          msg += "\n\n오류:\n" + errorMessages.slice(0, 5).join("\n");
          if (errorMessages.length > 5) {
            msg += `\n... 외 ${errorMessages.length - 5}개`;
          }
        }
        alert(msg);
      } else {
        alert(`${addedCount}명의 학생이 추가되었습니다.`);
      }
    } else {
      if (errorMessages.length > 0) {
        alert(errorMessages.slice(0, 10).join("\n"));
      } else {
        alert("추가된 학생이 없습니다.\n(중복 또는 빈 항목)");
      }
    }
  }

  // 학생 리스트 업데이트
  function updateStudentList() {
    studentList.innerHTML = "";

    studentsList.forEach((student, index) => {
      const certiNum = getCertiNum(index);
      const item = document.createElement("div");
      item.className = "student-item";
      item.innerHTML = `
                <div class="student-info">
                    <strong>${
                      index + 1
                    }. ${student}</strong> (일련번호: ${certiNum})
                </div>
                <button class="delete-btn" onclick="removeStudent(${index})">삭제</button>
            `;
      studentList.appendChild(item);
    });
  }

  // 학생 삭제 (전역 함수로 노출)
  window.removeStudent = function (index) {
    if (confirm(`'${studentsList[index]}'를 삭제하시겠습니까?`)) {
      studentsList.splice(index, 1);
      updateStudentList();
    }
  };

  // 일련번호 생성
  function getCertiNum(index) {
    const startValue = certiNumStart.value.trim();
    if (!startValue) {
      return `${index + 1}/${index + 1}`;
    }
    return formatCertiNum(startValue, index);
  }

  // 일련번호 포맷팅
  function formatCertiNum(startStr, index) {
    // 숫자 추출
    const numbers = startStr.match(/\d+/g);
    if (!numbers || numbers.length === 0) {
      return `${index + 1}/${index + 1}`;
    }

    // 마지막 숫자 추출
    const baseNum = parseInt(numbers[numbers.length - 1]);
    const currentNum = baseNum + index;

    // 마지막 숫자 찾아서 교체
    const lastNum = numbers[numbers.length - 1];
    const lastNumIndex = startStr.lastIndexOf(lastNum);

    if (lastNumIndex !== -1) {
      // 같은 자릿수로 포맷팅
      const formattedNum = String(currentNum).padStart(lastNum.length, "0");
      const result =
        startStr.substring(0, lastNumIndex) +
        formattedNum +
        startStr.substring(lastNumIndex + lastNum.length);
      return result;
    }

    return `${index + 1}/${index + 1}`;
  }

  // 엑셀로 출력
  function exportToExcel() {
    if (studentsList.length === 0) {
      alert("출력할 학생명이 없습니다.");
      return;
    }

    // SheetJS가 로드되었는지 확인
    if (typeof XLSX === "undefined") {
      alert(
        "Excel 라이브러리가 로드되지 않았습니다. 페이지를 새로고침해주세요."
      );
      return;
    }

    // 워크북 생성
    const wb = XLSX.utils.book_new();

    // 헤더
    const headers = [["일련번호", "학생이름", "과정명", "장소", "날짜"]];

    // 데이터
    const data = studentsList.map((student, index) => [
      getCertiNum(index),
      student,
      subName.value,
      location.value,
      eduDate.value,
    ]);

    // 워크시트 생성
    const ws = XLSX.utils.aoa_to_sheet([...headers, ...data]);
    XLSX.utils.book_append_sheet(wb, ws, "학생명단");

    // 파일 다운로드
    XLSX.writeFile(wb, "certificates.xlsx");
    alert("엑셀 파일이 생성되었습니다!");
  }

  // 워드로 출력
  async function exportToWord() {
    if (studentsList.length === 0) {
      alert("출력할 학생명이 없습니다.");
      return;
    }

    if (!templateFile) {
      alert("샘플 Word 파일을 먼저 선택해주세요.");
      return;
    }

    // JSZip이 로드되었는지 확인
    if (typeof JSZip === "undefined") {
      alert(
        "JSZip 라이브러리가 로드되지 않았습니다. 페이지를 새로고침해주세요."
      );
      return;
    }

    try {
      // 템플릿 파일을 ArrayBuffer로 읽기 (한 번만 읽기)
      const arrayBuffer = await templateFile.arrayBuffer();

      // 각 학생마다 파일 생성 (딜레이를 두고 순차적으로 다운로드)
      for (let idx = 0; idx < studentsList.length; idx++) {
        const student = studentsList[idx];
        const certiNum = getCertiNum(idx);

        // 템플릿 파일 복사 및 수정
        const modifiedDocx = await processTemplateFile(
          arrayBuffer,
          certiNum,
          student,
          subName.value,
          location.value,
          eduDate.value,
          idx + 1
        );

        // 파일 다운로드
        const blob = new Blob([modifiedDocx], {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });
        const fileName = sanitizeFilename(student) + ".docx";

        // FileSaver 사용
        if (typeof saveAs !== "undefined") {
          saveAs(blob, fileName);
        } else {
          // FileSaver가 없을 경우 대체 방법
          const url = URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url;
          a.download = fileName;
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url);
        }

        // 각 파일 다운로드 사이에 딜레이 추가 (브라우저가 여러 파일을 처리할 수 있도록)
        if (idx < studentsList.length - 1) {
          await new Promise((resolve) => setTimeout(resolve, 300));
        }
      }

      alert(`${studentsList.length}개의 워드 파일이 생성되었습니다!`);

      // 페이지 리로딩
      setTimeout(() => {
        location.reload();
      }, 1000);
    } catch (error) {
      console.error("Error:", error);
      alert(`워드 파일 생성 중 오류가 발생했습니다:\n${error.message}`);
    }
  }

  // 템플릿 파일 처리
  async function processTemplateFile(
    arrayBuffer,
    certiNum,
    studName,
    subName,
    location,
    eduDate,
    studentIndex
  ) {
    // JSZip으로 docx 파일 열기 (docx는 ZIP 형식)
    const zip = await JSZip.loadAsync(arrayBuffer);

    // word/document.xml 읽기
    const documentXml = await zip.file("word/document.xml").async("string");

    // 텍스트 교체
    let modifiedXml = documentXml;

    // 교체할 텍스트 매핑
    const replacements = {
      "Dayeon Mun": studName,
      "MDR/ISO13485:2016 Internal Auditor Training Course\n16Hours/2Days Training Course":
        subName,
      "MDR/ISO13485:2016 Internal Auditor Training Course": subName,
      "16Hours/2Days Training Course": subName,
      "Seoul, Korea": location,
      "11th – 12th December 2025": eduDate,
      "11th - 12th December 2025": eduDate,
    };

    // 텍스트 교체 (XML 내에서 직접 교체)
    Object.keys(replacements).forEach((oldText) => {
      const newText = replacements[oldText];
      // XML 내의 텍스트는 <w:t> 태그 안에 있음
      const escapedOld = escapeXml(oldText);
      const escapedNew = escapeXml(newText);
      // <w:t> 태그 내의 텍스트 교체
      modifiedXml = modifiedXml.replace(
        new RegExp(
          `(<w:t[^>]*>)([^<]*${escapeRegex(escapedOld)}[^<]*)(</w:t>)`,
          "g"
        ),
        `$1${escapedNew}$3`
      );
    });

    // 일련번호 패턴 교체
    const certiNumPattern = /26\/ISO13485:2016-IAC-IH-KR\/(\d+)/g;
    const certiNumReplacement = `26/ISO13485:2016-IAC-IH-KR/${String(
      studentIndex
    ).padStart(2, "0")}`;
    modifiedXml = modifiedXml.replace(certiNumPattern, certiNumReplacement);

    // 수정된 XML을 ZIP에 다시 쓰기
    zip.file("word/document.xml", modifiedXml);

    // ZIP을 Blob으로 변환
    const blob = await zip.generateAsync({ type: "blob" });
    return blob;
  }

  // XML 특수문자 이스케이프
  function escapeXml(text) {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  // 정규식 특수문자 이스케이프
  function escapeRegex(text) {
    return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  // 파일명 정리
  function sanitizeFilename(filename) {
    const invalidChars = /[<>:"/\\|?*]/g;
    return filename.replace(invalidChars, "_").replace(/\s/g, "_").trim();
  }
});
