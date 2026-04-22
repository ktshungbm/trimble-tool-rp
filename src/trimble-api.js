/**
 * trimble-api.js — Trimble Connect Workspace API Connection
 * Copied from DDC-Statistics architecture.
 */
import { connect } from "trimble-connect-workspace-api";

// ── State ──
let api = null;
let viewer = null;

// ── Public API ──
export function getApi() { return api; }
export function getViewer() { return viewer; }

// ── Main Initialization ──
export async function initTrimble() {
    console.log("[TC Extension] Initializing...");

    // Check if we're inside an iframe (required for TC extension)
    const isInIframe = window.parent !== window;
    if (!isInIframe) {
        console.warn("[TC Extension] Not running inside an iframe. TC Workspace API requires iframe context.");
        alert("Cảnh báo: Đang chạy ngoài Trimble Connect (không nằm trong iframe).");
        return null;
    }

    try {
        // Connect with event callback — THE CORRECT WAY for TC Workspace API
        api = await connect(
            window.parent,
            (event, data) => {
                console.log("[TC Event]", event, data);
            },
            15000 // 15 second timeout
        );

        viewer = api.viewer;
        console.log("[TC Extension] Connected to Workspace API successfully");
        console.log("[TC Extension] API keys:", Object.keys(api));

        // Request access token permission (often required by TC)
        try {
            await api.extension.requestPermission("accesstoken");
            console.log("[TC Extension] Access token permission granted");
        } catch (e) {
            console.warn("[TC Extension] Permission request failed:", e.message || e);
        }

        return api;
    } catch (error) {
        console.error("[TC Extension] Connection failed:", error);
        if (error.message) console.error("[TC Extension] Lỗi:", error.message);
        alert("[TC Extension] Kết nối API thất bại: " + error);
        return null;
    }
}

// ── Utils ──
function chunkArray(arr, size) {
    const out = [];
    for (let i = 0; i < arr.length; i += size) {
        out.push(arr.slice(i, i + size));
    }
    return out;
}

function extractRuntimeIdsFromGroup(group) {
    const ids = new Set();
    if (!group) return [];
    if (Array.isArray(group.objects)) {
        for (const obj of group.objects) {
            if (typeof obj?.id === "number") ids.add(obj.id);
        }
    }
    return Array.from(ids);
}

export async function grayOutModel() {
    // Để giữ tương thích ngược nếu lỡ gọi từ ngoài vào, ta không xử lý gì ở đây nữa 
    // vì ta sẽ phủ trắng bằng cách batch process trực tiếp bên trong processModelObjects
}

function normalizeConvertedIds(value) {
    if (value === null || value === undefined) return [];
    if (typeof value === "number") return [value];
    if (Array.isArray(value)) return value.flat(Infinity).filter(v => typeof v === "number");
    return [];
}

async function setObjectColorBatch(modelId, runtimeIds, colorString) {
    const batches = chunkArray(runtimeIds, 1000);
    for (const ids of batches) {
        if (!ids.length) continue;
        await viewer.setObjectState(
            { modelObjectIds: [{ modelId: modelId, objectRuntimeIds: ids }] },
            { color: colorString }
        );
    }
}

export async function processModelObjects(issuedAssemblies, btnIssue, renderReportCallback) {
    if (!api || !viewer) {
        if (btnIssue) btnIssue.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i> Lỗi API';
        return;
    }

    try {
        let reportData = [];
        if (btnIssue) btnIssue.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang tải cấu trúc từ Trimble...';

        const rawObjects = await viewer.getObjects();
        if (!Array.isArray(rawObjects) || rawObjects.length === 0) {
            if (btnIssue) btnIssue.innerHTML = '<i class="fa-solid fa-paintbrush"></i> Mô hình trống';
            return;
        }

        const modelGroups = rawObjects.map(group => ({
            modelId: group?.modelId,
            runtimeIds: extractRuntimeIdsFromGroup(group)
        })).filter(g => g.modelId && g.runtimeIds.length > 0);

        if (!modelGroups.length) {
            if (btnIssue) btnIssue.innerHTML = '<i class="fa-solid fa-triangle-exclamation"></i> Không thể nạp Runtime IDs';
            return;
        }

        // 1. Phủ Trắng mô hình (#FFFFFF) chạy theo từng model
        if (btnIssue) btnIssue.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang làm trắng mô hình...';
        for (const group of modelGroups) {
            await setObjectColorBatch(group.modelId, group.runtimeIds, "#FFFFFF");
        }

        // 2. Định danh GUIDs (Ban hành)
        if (btnIssue) btnIssue.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang trích xuất định danh...';
        let idsToConvert = issuedAssemblies.map(a => typeof a === 'object' ? a.id : a);

        for (const group of modelGroups) {
            const modelId = group.modelId;
            let convertedIds = [];
            
            try {
                let runtimeIdsResponse = await viewer.convertToObjectRuntimeIds(modelId, idsToConvert);
                if (Array.isArray(runtimeIdsResponse)) {
                    for (let i = 0; i < issuedAssemblies.length; i++) {
                        const ids = normalizeConvertedIds(runtimeIdsResponse[i]);
                        const label = typeof issuedAssemblies[i] === 'object' ? issuedAssemblies[i].label : issuedAssemblies[i];
                        if (ids.length > 0) {
                            convertedIds.push(...ids);
                            // Report updates
                            reportData.push({ mark: label, issued: true, modelId: modelId, objectRuntimeIds: ids });
                        } else if (!reportData.find(r => r.mark === label)) {
                            reportData.push({ mark: label, issued: false });
                        }
                    }
                }
            } catch (err) {
                console.error(`Lỗi chuyển đổi GUID trên model ${modelId}:`, err);
            }

            // 3. Tô màu Xanh lá (#20C000) những cấu kiện ban hành
            if (convertedIds.length > 0) {
                if (btnIssue) btnIssue.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang đắp màu xanh...';
                
                // Select trên Viewer (chỉ batch giới hạn)
                const selectionBatches = chunkArray(convertedIds, 1000);
                for (const ids of selectionBatches) {
                     await viewer.setSelection({ modelObjectIds: [{ modelId: modelId, objectRuntimeIds: ids }] }, "set");
                }
                
                // Auto Tô Màu
                await setObjectColorBatch(modelId, convertedIds, "#20C000");
            }
        }

        // 5. Báo cáo (loại bỏ trùng lặp)
        const uniqueReportItems = Array.from(new Map(reportData.map(item => [item.mark, item])).values());
        if (renderReportCallback) renderReportCallback(uniqueReportItems);

        // 6. Tự động lưu View Ban Hành (và giữ tối đa 5 view)
        if (btnIssue) btnIssue.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Đang lưu View...';
        await autoSaveIssueView();

        // Hoàn thành
        if (btnIssue) {
            btnIssue.classList.replace('bg-blue-600', 'bg-emerald-500');
            btnIssue.classList.replace('hover:bg-blue-700', 'hover:bg-emerald-600');
            btnIssue.innerHTML = '<i class="fa-solid fa-check"></i> Đã tô màu, Báo Cáo & Lưu View!';

            setTimeout(() => {
                btnIssue.classList.add('hidden');
                btnIssue.classList.replace('bg-emerald-500', 'bg-blue-600');
                btnIssue.classList.replace('hover:bg-emerald-600', 'hover:bg-blue-700');
            }, 4000);
        }

    } catch (error) {
        console.error("Lỗi API:", error);
        let errorMsg = error.message ? error.message.substring(0, 30) : "Unknown Error";
        if (btnIssue) btnIssue.innerHTML = `<i class="fa-solid fa-xmark"></i> ${errorMsg}`;
    }
}

// ── Tự động lưu View và giữ tối đa 5 View Ban Hành ──
async function autoSaveIssueView() {
    if (!api || !api.view) return;
    try {
        let viewName = "Ban Hành - " + new Date().toLocaleString('vi-VN');
        await api.view.createView({ name: viewName, description: "Tự động lưu trạng thái ban hành bởi BIM STEEL" });
        console.log("[TC Extension] Đã lưu view:", viewName);

        let views = await api.view.getViews();
        let issueViews = views.filter(v => v.name && v.name.startsWith("Ban Hành - ")).sort((a, b) => {
            return new Date(a.createdOn).getTime() - new Date(b.createdOn).getTime(); // cũ nhất lên đầu
        });

        if (issueViews.length > 5) {
            let excess = issueViews.length - 5;
            for (let i = 0; i < excess; i++) {
                await api.view.deleteView(issueViews[i].id);
                console.log("[TC Extension] Đã xóa view cũ:", issueViews[i].name);
            }
        }
    } catch (err) {
        console.error("[TC Extension] Lỗi khi lưu view ban hành:", err);
    }
}

export async function loadLatestIssueView() {
    if (!api || !api.view) {
        return grayOutModel();
    }
    try {
        let views = await api.view.getViews();
        let issueViews = views.filter(v => v.name && v.name.startsWith("Ban Hành - ")).sort((a, b) => {
            return new Date(b.createdOn).getTime() - new Date(a.createdOn).getTime(); // mới nhất lên đầu
        });

        if (issueViews.length > 0) {
            let latestView = issueViews[0];
            console.log("[TC Extension] Đang load view mới nhất:", latestView.name);
            await api.view.selectView(latestView.id);
        } else {
            console.log("[TC Extension] Không có view ban hành nào, tiến hành phủ xám.");
            await grayOutModel();
        }
    } catch (err) {
        console.error("[TC Extension] Lỗi khi load view mới nhất:", err);
        await grayOutModel();
    }
}

export async function selectModelObject(modelId, objectRuntimeIds) {
    if (!api || !viewer) return;
    try {
        await viewer.setSelection({
            modelObjectIds: [{ modelId: modelId, objectRuntimeIds: objectRuntimeIds }]
        }, "set");
    } catch (err) {
        console.error("[TC Extension] Lỗi khi select đối tượng:", err);
    }
}
