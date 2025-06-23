document.addEventListener('DOMContentLoaded', function() {
    const dropZone = document.getElementById('dropZone');
    const fileInfo = document.getElementById('fileInfo');
    const validationContainer = document.getElementById('validationContainer');
    const sheet1Container = document.getElementById('sheet1Container');
    const sheet2Container = document.getElementById('sheet2Container');
    const sheet1Table = document.getElementById('sheet1Table');
    const sheet2Table = document.getElementById('sheet2Table');
    
    let workbook = null;
    let sheetData = {};
    let currentFileName = ''; // 現在のファイル名を保存する変数

    // ドラッグオーバーイベント
    dropZone.addEventListener('dragover', function(e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add('active');
    });

    // ドラッグリーブイベント
    dropZone.addEventListener('dragleave', function(e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('active');
    });

    // ドロップイベント
    dropZone.addEventListener('drop', function(e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove('active');
        
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            const file = files[0];
            if (isExcelFile(file)) {
                processExcelFile(file);
            } else {
                showError('Excel ファイル（.xlsx または .xls）をドロップしてください。');
            }
        }
    });

    // Excelファイルかどうかをチェック
    function isExcelFile(file) {
        return file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
    }

    // エラーメッセージを表示
    function showError(message) {
        validationContainer.innerHTML = `<div class="error-message">${message}</div>`;
        validationContainer.classList.remove('hidden');
    }

    // 警告メッセージを表示
    function showWarning(message) {
        validationContainer.innerHTML += `<div class="warning-message">${message}</div>`;
        validationContainer.classList.remove('hidden');
    }

    // バリデーションメッセージをクリア
    function clearValidationMessages() {
        validationContainer.innerHTML = '';
        validationContainer.classList.add('hidden');
    }
    
    // ExcelSTLのデータ構造を検証
    function validateExcelSTLStructure() {
        clearValidationMessages();
        
        let hasErrors = false;
        let mode = '';
        
        // シート存在チェック
        if (!workbook.SheetNames.includes('Sheet1')) {
            showError('Sheet1が見つかりません。頂点座標はSheet1に記載してください。');
            hasErrors = true;
        }
        
        if (!workbook.SheetNames.includes('Sheet2')) {
            showError('Sheet2が見つかりません。三角形インデックスリストはSheet2に記載してください。');
            hasErrors = true;
        }
        
        // Sheet1のデータを取得
        const sheet1Data = sheetData['Sheet1'];
        if (!sheet1Data || sheet1Data.length === 0) {
            showError('Sheet1にデータがありません。');
            hasErrors = true;
        }
        
        // Sheet1の列数チェック（存在する場合のみ）
        if (sheet1Data && sheet1Data.length > 0) {
            const nonEmptyRows = sheet1Data.filter(row => row && row.length > 0);
            if (nonEmptyRows.length > 0) {
                const sheet1Cols = Math.max(...nonEmptyRows.map(row => row.length));
                
                if (sheet1Cols === 2) {
                    mode = '2D';
                } else if (sheet1Cols === 3) {
                    mode = '3D';
                } else {
                    showError(`Sheet1の列数が不正です。2D モードでは2列、3D モードでは3列のデータが必要です。現在: ${sheet1Cols}列`);
                    hasErrors = true;
                }
            }
        }
        
        // Sheet2のデータを取得
        const sheet2Data = sheetData['Sheet2'];
        if (!sheet2Data || sheet2Data.length === 0) {
            showError('Sheet2にデータがありません。');
            hasErrors = true;
        }
        
        // Sheet2の列数チェック（存在する場合のみ）
        if (sheet2Data && sheet2Data.length > 0) {
            const nonEmptyRows = sheet2Data.filter(row => row && row.length > 0);
            if (nonEmptyRows.length > 0) {
                const sheet2Cols = Math.max(...nonEmptyRows.map(row => row.length));
                
                if (sheet2Cols !== 3) {
                    showError(`Sheet2の列数が不正です。三角形インデックスは3列で記述してください。現在: ${sheet2Cols}列`);
                    hasErrors = true;
                }
            }
        }
        
        // Sheet1とSheet2の両方が存在する場合のみインデックスをチェック
        if (sheet1Data && sheet1Data.length > 0 && sheet2Data && sheet2Data.length > 0) {
            // Sheet1の頂点数
            const vertexCount = sheet1Data.length;
            
            // Sheet2のインデックスチェック
            let hasInvalidIndices = false;
            let indicesTooLarge = false;
            let indicesNegativeOrZero = false;
            
            sheet2Data.forEach((row, rowIndex) => {
                if (!row) return;
                
                for (let i = 0; i < 3 && i < row.length; i++) {
                    const index = row[i];
                    
                    // 数値チェック
                    if (typeof index !== 'number') {
                        hasInvalidIndices = true;
                    } 
                    // 範囲チェック（1からvertexCount）
                    else if (index > vertexCount) {
                        indicesTooLarge = true;
                    } 
                    // 0以下のチェック
                    else if (index <= 0) {
                        indicesNegativeOrZero = true;
                    }
                }
            });
            
            if (hasInvalidIndices) {
                showError('Sheet2に無効なインデックスがあります。全てのセルが数値である必要があります。');
                hasErrors = true;
            }
            
            if (indicesTooLarge) {
                showError(`Sheet2のインデックスが頂点数を超えています。インデックスは1から${vertexCount}の範囲内である必要があります。`);
                hasErrors = true;
            }
            
            if (indicesNegativeOrZero) {
                showError('Sheet2にゼロ以下のインデックスがあります。インデックスは1以上である必要があります。');
                hasErrors = true;
            }
        }
        
        // 成功メッセージは、エラーがない場合のみ表示
        if (!hasErrors && mode) {
            const vertexCount = (sheetData['Sheet1'] || []).length;
            const triangleCount = (sheetData['Sheet2'] || []).length;
            showWarning(`検証完了: ${mode}モードとして認識されました。Sheet1に${vertexCount}個の頂点、Sheet2に${triangleCount}個の三角形が定義されています。`);
        }
        
        return !hasErrors;
    }
      // Excelファイルを処理
    function processExcelFile(file) {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                workbook = XLSX.read(data, { type: 'array' });
                
                // ファイル名を保存（拡張子なし）
                currentFileName = file.name.replace(/\.(xlsx|xls)$/i, '');
                
                // シートデータをキャッシュ
                sheetData = {};
                workbook.SheetNames.forEach(name => {
                    const worksheet = workbook.Sheets[name];
                    sheetData[name] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                });
                
                // ファイル情報を表示
                fileInfo.textContent = `ファイル: ${file.name}`;
                fileInfo.classList.remove('hidden');
                
                // ExcelSTLの構造を検証
                validateExcelSTLStructure();
                
                // STL生成ボタンを表示
                document.getElementById('stlGenerationContainer').classList.remove('hidden');
                
                // Sheet1とSheet2を表示（エラーがあっても表示する）
                if (sheetData['Sheet1']) {
                    renderSheet1(sheetData['Sheet1']);
                }
                
                if (sheetData['Sheet2']) {
                    renderSheet2(sheetData['Sheet2']);
                }
            } catch (error) {
                console.error('ファイル処理エラー:', error);
                showError('ファイルの読み込み中にエラーが発生しました: ' + error.message);
            }
        };
        
        reader.onerror = function() {
            showError('ファイルの読み込みに失敗しました。');
        };
        
        reader.readAsArrayBuffer(file);
    }
    
    // Sheet1を表示 (頂点座標)
    function renderSheet1(data) {
        sheet1Table.innerHTML = '';
        
        if (!data || data.length === 0) {
            sheet1Table.textContent = 'データが空です。';
            sheet1Container.classList.remove('hidden');
            return;
        }
        
        const table = document.createElement('table');
        const maxCols = Math.max(...data.map(row => row.length || 0));
        
        // 2D/3Dに応じた列数処理（デフォルトは2D）
        const columnCount = maxCols > 0 ? Math.min(maxCols, 3) : 2;
        
        // ヘッダー行を生成
        const headerRow = document.createElement('tr');
        
        // 行番号用のヘッダーセル
        const rowNumHeader = document.createElement('th');
        rowNumHeader.textContent = '頂点番号';
        headerRow.appendChild(rowNumHeader);
        
        // 列ヘッダー
        const headers = ['x', 'y', 'z'];
        for (let i = 0; i < columnCount; i++) {
            const th = document.createElement('th');
            th.textContent = headers[i];
            headerRow.appendChild(th);
        }
        
        table.appendChild(headerRow);
        
        // データ行を生成
        data.forEach((row, rowIndex) => {
            const tr = document.createElement('tr');
            
            // 行番号のセル (頂点番号)
            const rowNumCell = document.createElement('td');
            rowNumCell.textContent = rowIndex + 1;
            rowNumCell.style.backgroundColor = '#f2f2f2';
            rowNumCell.style.fontWeight = 'bold';
            tr.appendChild(rowNumCell);
            
            // データセル
            for (let i = 0; i < columnCount; i++) {
                const td = document.createElement('td');
                const cellValue = row && i < row.length ? row[i] : '';
                td.textContent = cellValue;
                tr.appendChild(td);
            }
            
            table.appendChild(tr);
        });
        
        sheet1Table.appendChild(table);
        sheet1Container.classList.remove('hidden');
    }

    // Sheet2を表示 (三角形インデックスリスト)
    function renderSheet2(data) {
        sheet2Table.innerHTML = '';
        
        if (!data || data.length === 0) {
            sheet2Table.textContent = 'データが空です。';
            sheet2Container.classList.remove('hidden');
            return;
        }
        
        const table = document.createElement('table');
        
        // ヘッダー行を生成
        const headerRow = document.createElement('tr');
        
        // 行番号用のヘッダーセル
        const rowNumHeader = document.createElement('th');
        rowNumHeader.textContent = '三角形番号';
        headerRow.appendChild(rowNumHeader);
        
        // 列ヘッダー
        const headers = ['v1', 'v2', 'v3'];
        for (let i = 0; i < 3; i++) {
            const th = document.createElement('th');
            th.textContent = headers[i];
            headerRow.appendChild(th);
        }
        
        table.appendChild(headerRow);
        
        // データ行を生成
        data.forEach((row, rowIndex) => {
            const tr = document.createElement('tr');
            
            // 行番号のセル (三角形番号)
            const rowNumCell = document.createElement('td');
            rowNumCell.textContent = rowIndex + 1;
            rowNumCell.style.backgroundColor = '#f2f2f2';
            rowNumCell.style.fontWeight = 'bold';
            tr.appendChild(rowNumCell);
            
            // データセル
            for (let i = 0; i < 3; i++) {
                const td = document.createElement('td');
                const cellValue = row && i < row.length ? row[i] : '';
                td.textContent = cellValue;
                
                // エラーハイライト
                const vertexCount = (sheetData['Sheet1'] || []).length;
                if (typeof cellValue === 'number') {
                    if (cellValue > vertexCount) {
                        td.style.backgroundColor = '#ffebee';
                        td.style.color = '#d32f2f';
                        td.title = `インデックスが頂点数(${vertexCount})を超えています`;
                    } else if (cellValue <= 0) {
                        td.style.backgroundColor = '#ffebee';
                        td.style.color = '#d32f2f';
                        td.title = 'インデックスは1以上である必要があります';
                    }
                }
                
                tr.appendChild(td);
            }
            
            table.appendChild(tr);
        });
        
        sheet2Table.appendChild(table);
        sheet2Container.classList.remove('hidden');
    }

    // STLファイルを生成する関数
    function generateSTL() {
        if (!sheetData['Sheet1'] || !sheetData['Sheet2']) {
            showError('有効なデータがありません。');
            return;
        }
        
        try {
            // 頂点データと三角形インデックスを取得
            const vertices = sheetData['Sheet1'];
            const indices = sheetData['Sheet2'];
            
            if (vertices.length === 0 || indices.length === 0) {
                showError('有効なデータがありません。');
                return;
            }

            // 単位を取得
            const unitElements = document.getElementsByName('unit');
            let dataUnit = 'cm'; // デフォルト
            for (const element of unitElements) {
                if (element.checked) {
                    dataUnit = element.value;
                    break;
                }
            }
            
            // 押し出し厚さを取得
            const thickness = parseFloat(document.getElementById('thickness').value) || 3;
            
            // 2D/3Dモードを判定
            const is2DMode = vertices[0] && vertices[0].length <= 2;
            
            // STL文字列の生成
            let stlString;
            if (is2DMode) {
                // 2Dモード (押し出し処理)
                stlString = generateSTLFrom2D(vertices, indices, dataUnit, thickness);
            } else {
                // 3Dモード
                stlString = generateSTLFrom3D(vertices, indices, dataUnit);
            }
            
            // STLファイルをダウンロード
            downloadSTL(stlString);
            
            showWarning('STLファイルの生成が完了しました。');
        } catch (error) {
            console.error('STL生成エラー:', error);
            showError('STLファイルの生成中にエラーが発生しました: ' + error.message);
        }
    }
    
    // 2Dデータから押し出しでSTLを生成
    function generateSTLFrom2D(vertices, indices, unit, thickness) {
        // 単位変換 (cmからmmへの変換係数)
        const scaleFactor = unit === 'cm' ? 10 : 1;
        
        // 正規化された頂点データを作成
        const normalizedVertices = vertices.map(v => {
            // 2次元データを3次元に変換して単位も変換
            return [
                (v[0] || 0) * scaleFactor,
                (v[1] || 0) * scaleFactor,
                0 // Z座標は0
            ];
        });
        
        // 三角形インデックスの向きを修正
        const fixedIndices = fixTriangleDirection(normalizedVertices, indices);
        
        // 押し出し処理
        const extrudedMesh = extrudeMesh(normalizedVertices, fixedIndices, thickness);
        
        // STL文字列を生成
        return makeSTLString(extrudedMesh.vertices, extrudedMesh.indices);
    }
    
    // 3DデータからSTLを生成
    function generateSTLFrom3D(vertices, indices, unit) {
        // 単位変換 (cmからmmへの変換係数)
        const scaleFactor = unit === 'cm' ? 10 : 1;
        
        // 正規化された頂点データを作成
        const normalizedVertices = vertices.map(v => {
            return [
                (v[0] || 0) * scaleFactor,
                (v[1] || 0) * scaleFactor,
                (v[2] || 0) * scaleFactor
            ];
        });
        
        // STL文字列を生成
        return makeSTLString(normalizedVertices, indices);
    }
    
    // 三角形の頂点の向きを修正（反時計回りにする）
    function fixTriangleDirection(vertices, indices) {
        const result = indices.map(triangle => {
            // インデックスは1-basedなので0-basedに変換
            const v1Idx = triangle[0] - 1;
            const v2Idx = triangle[1] - 1;
            const v3Idx = triangle[2] - 1;
            
            if (v1Idx < 0 || v2Idx < 0 || v3Idx < 0 || 
                v1Idx >= vertices.length || v2Idx >= vertices.length || v3Idx >= vertices.length) {
                return triangle; // 無効なインデックスはそのまま返す
            }
            
            const v1 = vertices[v1Idx];
            const v2 = vertices[v2Idx];
            const v3 = vertices[v3Idx];
            
            // 2Dの場合、Z座標は無視して外積を計算
            const a = [v2[0] - v1[0], v2[1] - v1[1]];
            const b = [v3[0] - v1[0], v3[1] - v1[1]];
            const crossProduct = a[0] * b[1] - a[1] * b[0];
            
            // 外積が負の場合、v2とv3を入れ替え
            if (crossProduct < 0) {
                return [triangle[0], triangle[2], triangle[1]];
            }
            
            return triangle;
        });
        
        return result;
    }
    
    // 2Dメッシュを押し出して3Dメッシュを作成
    function extrudeMesh(vertices, indices, thickness) {
        const vertexCount = vertices.length;
        const triangleCount = indices.length;
        
        // 新しい頂点配列（上面と下面）
        let newVertices = [];
        
        // 上面の頂点を追加
        for (const v of vertices) {
            newVertices.push([v[0], v[1], thickness]);
        }
        
        // 下面の頂点を追加
        for (const v of vertices) {
            newVertices.push([v[0], v[1], 0]);
        }
        
        // 新しい三角形インデックス配列
        let newIndices = [];
        
        // 上面の三角形を追加
        for (const triangle of indices) {
            newIndices.push(triangle);
        }
        
        // 下面の三角形を追加（頂点の順序を反転して下面を表す）
        for (const triangle of indices) {
            const bottomTriangle = [
                triangle[0] + vertexCount,
                triangle[2] + vertexCount,
                triangle[1] + vertexCount
            ];
            newIndices.push(bottomTriangle);
        }
        
        // 側面の三角形を追加
        for (const triangle of indices) {
            for (let i = 0; i < 3; i++) {
                const v1 = triangle[i];
                const v2 = triangle[(i+1) % 3];
                const v1Bottom = v1 + vertexCount;
                const v2Bottom = v2 + vertexCount;
                
                // 側面の四角形を2つの三角形に分割
                newIndices.push([v1, v2, v1Bottom]);
                newIndices.push([v2, v2Bottom, v1Bottom]);
            }
        }
        
        return {
            vertices: newVertices,
            indices: newIndices
        };
    }
    
    // STL文字列を生成
    function makeSTLString(vertices, indices) {
        let stl = 'solid ExcelSTL\n';
        
        indices.forEach(triangle => {
            // インデックスは1-basedなので0-basedに変換
            const v1Idx = Math.max(0, triangle[0] - 1);
            const v2Idx = Math.max(0, triangle[1] - 1);
            const v3Idx = Math.max(0, triangle[2] - 1);
            
            if (v1Idx >= vertices.length || v2Idx >= vertices.length || v3Idx >= vertices.length) {
                return; // 無効なインデックスはスキップ
            }
            
            const v1 = vertices[v1Idx];
            const v2 = vertices[v2Idx];
            const v3 = vertices[v3Idx];
            
            // 法線ベクトルを計算
            const normal = calculateNormal(v1, v2, v3);
            
            // 三角形を追加
            stl += '  facet normal ' + normal.join(' ') + '\n';
            stl += '    outer loop\n';
            stl += '      vertex ' + v1.join(' ') + '\n';
            stl += '      vertex ' + v2.join(' ') + '\n';
            stl += '      vertex ' + v3.join(' ') + '\n';
            stl += '    endloop\n';
            stl += '  endfacet\n';
        });
        
        stl += 'endsolid ExcelSTL\n';
        return stl;
    }
    
    // 法線ベクトルを計算
    function calculateNormal(v1, v2, v3) {
        // 辺ベクトルを計算
        const edge1 = [v2[0] - v1[0], v2[1] - v1[1], v2[2] - v1[2]];
        const edge2 = [v3[0] - v1[0], v3[1] - v1[1], v3[2] - v1[2]];
        
        // 外積を計算
        const normal = [
            edge1[1] * edge2[2] - edge1[2] * edge2[1],
            edge1[2] * edge2[0] - edge1[0] * edge2[2],
            edge1[0] * edge2[1] - edge1[1] * edge2[0]
        ];
        
        // 正規化
        const length = Math.sqrt(normal[0] * normal[0] + normal[1] * normal[1] + normal[2] * normal[2]);
        if (length > 0) {
            normal[0] /= length;
            normal[1] /= length;
            normal[2] /= length;
        }
        
        return normal;
    }
      // STLファイルをダウンロード
    function downloadSTL(stlString) {
        const blob = new Blob([stlString], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        
        // 入力したExcelファイル名に基づいてSTLファイル名を設定
        // デフォルト名はmodel.stlとする
        const stlFileName = currentFileName ? `${currentFileName}.stl` : 'model.stl';
        a.download = stlFileName;
        
        document.body.appendChild(a);
        a.click();
        
        // クリーンアップ
        setTimeout(function() {
            document.body.removeChild(a);
            window.URL.revokeObjectURL(url);
        }, 0);
    }
    
    // 2D/3Dモードに応じて押し出し厚さのコントロールを表示/非表示
    function updateThicknessVisibility() {
        const thicknessGroup = document.getElementById('thicknessGroup');
        if (sheetData['Sheet1'] && sheetData['Sheet1'][0]) {
            const is2DMode = sheetData['Sheet1'][0].length <= 2;
            thicknessGroup.style.display = is2DMode ? 'block' : 'none';
        }
    }

    // STL生成ボタンのイベントハンドラを設定
    document.getElementById('generateStlBtn').addEventListener('click', generateSTL);
    console.log('STL生成ボタンにイベントハンドラを設定しました');
    
    // ファイルが読み込まれたら押し出し厚さの表示を更新
    const observer = new MutationObserver(function(mutations) {
        mutations.forEach(function(mutation) {
            if (mutation.type === 'attributes' && mutation.attributeName === 'class') {
                if (!sheet1Container.classList.contains('hidden')) {
                    updateThicknessVisibility();
                }
            }
        });
    });
    
    observer.observe(sheet1Container, { attributes: true });
});
