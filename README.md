# CSB Inventory

Static site for the Center for Symbiotic Biomaterials (CSB) equipment inventory dashboard.

## 🚀 Deployment

The site is published via GitHub Pages from the `main` branch using the GitHub Actions workflow in `.github/workflows/deploy.yml`.

---

## 📝 Collaborator Guide (협업자를 위한 가이드)

This repository is designed to be updated easily by non-technical collaborators working out of the shared Dropbox team folder.

공유 드롭박스 폴더를 사용하는 비개발자 협업자분들도 쉽게 대시보드 데이터를 업데이트하고 사진을 등록할 수 있도록 구성되어 있습니다.

### 🇰🇷 한국어 설명 (Collaborators용)

#### 1단계: 장비 정보 수정하기 (위치, 담당자, 비고)
1. 공유 폴더에 있는 **`담당자.xlsx`** 엑셀 파일을 엽니다.
2. **`장비 목록 (전체)`** 시트로 이동합니다.
3. 수정하고 싶은 장비의 **장비 위치, 담당자, 비고** 칸을 수정하고 저장 후 닫습니다.

#### 2단계: 신규 사진 추가하기
1. 등록할 사진 파일의 이름을 장비 번호에 맞춰 수정합니다. (예: 37번 장비 사진 -> **`equipment_37.jpeg`**)
   * *주의: 파일 형식은 반드시 `.jpeg` 또는 `.jpg`여야 하며 소문자로 입력해야 합니다.*
2. 이름 변경한 사진 파일을 **`images/`** 폴더 안에 넣어줍니다.

#### 3단계: 대시보드 반영하기 (가장 중요)
1. 폴더 내에 있는 **`update_inventory.bat`** 파일을 마우스로 더블 클릭합니다.
2. 검은색 창이 실행되며 자동으로 데이터를 변환하고 무결성 체크를 진행합니다.
3. 창에 **`[SUCCESS] Dashboard has been successfully updated!`** 메시지가 뜨면 아무 키나 눌러 창을 닫습니다.

#### 4단계: 확인하기
* **`index.html`** 파일을 더블 클릭하여 인터넷 브라우저로 띄운 뒤, 내가 수정한 정보와 사진이 올바르게 나타나는지 확인합니다.

---

### 🇺🇸 English Instructions

#### Step 1: Edit Equipment Details (Locations, Managers, Remarks)
1. Open the **`담당자.xlsx`** Excel file in the shared folder.
2. Go to the **`장비 목록 (전체)`** sheet.
3. Edit the **장비 위치 (Location)**, **담당자 (Manager)**, or **비고 (Remarks)** columns as needed. Save and close Excel.

#### Step 2: Add New Pictures
1. Rename your new equipment photo to match the equipment number (e.g., photo for equipment #37 -> **`equipment_37.jpeg`**).
   * *Note: The format must be `.jpeg` or `.jpg` in lowercase.*
2. Move the renamed photo into the **`images/`** directory.

#### Step 3: Apply Changes to Dashboard
1. Double-click the **`update_inventory.bat`** file in the main folder.
2. A command window will run the automated script to compile the new list.
3. Once you see **`[SUCCESS] Dashboard has been successfully updated!`**, press any key to close the window.

#### Step 4: Verify Changes
* Double-click and open **`index.html`** in your browser to verify that your new details and photos are loaded successfully.
