// ===== GLOBAL CONFIG =====
const SPREADSHEET_ID = "1xeNVv2yLADoxuZQwYBEZNvlZnt9CKvJ40RLTwNOrQfU";

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getPendingDeliverySheet() {
  const sheet =
    getSS().getSheetByName("pending_delivery");

  if (!sheet) {
    throw new Error(
      "pending_delivery sheet not found"
    );
  }

  return sheet;
}

function createPendingDelivery(data, by) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
  const sheet = getPendingDeliverySheet();

  /* ================= VALIDATE ================= */
  if (!data) {
    throw new Error("Missing pending delivery data");
  }

  const orderId =
    String(data.orderId || "").trim();

  const poNumber =
    String(data.poNumber || "").trim();

  const productId =
    String(data.productId || "").trim();

  const qty =
    Number(data.qty);

  const note =
    String(data.note || "").trim();

  if (
    !orderId ||
    !poNumber ||
    !productId ||
    !Number.isInteger(qty) ||
    qty <= 0
  ) {
    throw new Error(
      "Invalid pending delivery data"
    );
  }

  /* ================= VERIFY ORDER ================= */
  const orderSheet = getSS().getSheetByName("Orders");

  if (!orderSheet) {
    throw new Error("Orders sheet not found");
  }

  const orderRows = orderSheet.getDataRange().getValues();
  orderRows.shift();

  const orderRow = orderRows.find(
    row => String(row[0]).trim() === orderId
  );

  if (!orderRow) {
    throw new Error("Order not found");
  }

  const orderStatus =
    String(orderRow[3] || "").trim().toUpperCase();

  if (orderStatus !== "APPROVED") {
    throw new Error("Order is not approved");
  }

  const orderPoNumber =
    String(orderRow[7] || "").trim();

  if (orderPoNumber !== poNumber) {
    throw new Error("PO number does not match order");
  }

  const orderItems = JSON.parse(orderRow[1] || "[]");

  if (!Array.isArray(orderItems)) {
    throw new Error("Invalid order items");
  }

  const orderItem = orderItems.find(
    item => String(item.productId || "").trim() === productId
  );

  if (!orderItem) {
    throw new Error("Product not found in order");
  }

  const orderedQty = Number(orderItem.qty);

  if (
    !Number.isInteger(orderedQty) ||
    orderedQty <= 0
  ) {
    throw new Error("Invalid ordered quantity");
  }

  if (qty > orderedQty) {
    throw new Error("Pending quantity exceeds ordered quantity");
  }

 /* ================= DUPLICATE CHECK ================= */

  const rows =
    sheet.getDataRange().getValues();

  const headers =
    rows.shift();

  const orderIdx =
    headers.indexOf("orderId");

  const productIdx =
    headers.indexOf("productId");

  const statusIdx =
    headers.indexOf("status");

  if (
    orderIdx === -1 ||
    productIdx === -1 ||
    statusIdx === -1
  ) {
    throw new Error(
      "pending_delivery schema mismatch"
    );
  }

  const exists =
    rows.some(row =>
      String(row[orderIdx]).trim() === orderId &&
      String(row[productIdx]).trim() === productId &&
      String(row[statusIdx]).trim().toUpperCase() === "OPEN"
    );

  if (exists) {
    throw new Error(
      "รายการค้างส่งนี้มีอยู่แล้ว"
    );
  }  

  /* ================= CREATE ================= */
  const pendingId =
    "PD-" + Date.now();

  const createdAt =
    new Date();

  sheet.appendRow([
    pendingId,     // pendingId
    orderId,       // orderId
    poNumber,      // poNumber
    productId,     // productId
    orderedQty,
    qty,           // qty
    note,          // note
    "OPEN",        // status
    createdAt,     // createdAt
    by || "",      // createdBy
    "",            // closedAt
    ""             // closedBy
  ]);

  return {
    success: true,
    pendingId
  };
  } finally {
    lock.releaseLock();
  }
}

function getPendingDeliveries() {

  const sheet = getPendingDeliverySheet();

  const rows = sheet.getDataRange().getValues();

  if (rows.length < 2) {
    return [];
  }

  const headers =
    rows.shift().map(h => String(h || "").trim());

  return rows
    .map(row => {
      const obj = {};

      headers.forEach((header, index) => {
        obj[header] = row[index];
      });

      return obj;
    })
    .filter(item => item.pendingId);
}

function closePendingDelivery(pendingId, by) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
  const sheet = getPendingDeliverySheet();

  const rows = sheet.getDataRange().getValues();

  if (rows.length < 2) {
    throw new Error("No pending deliveries found");
  }

  const headers =
    rows[0].map(h => String(h || "").trim());

  const pendingIdCol =
    headers.indexOf("pendingId");

  const statusCol =
    headers.indexOf("status");

  const closedAtCol =
    headers.indexOf("closedAt");

  const closedByCol =
    headers.indexOf("closedBy");

  if (
    pendingIdCol === -1 ||
    statusCol === -1 ||
    closedAtCol === -1 ||
    closedByCol === -1
  ) {
    throw new Error(
      "pending_delivery schema mismatch"
    );
  }

  const rowIndex = rows.findIndex(
    (row, index) =>
      index > 0 &&
      String(row[pendingIdCol]).trim() ===
      String(pendingId).trim()
  );

  if (rowIndex === -1) {
    throw new Error("Pending delivery not found");
  }

  const currentStatus =
    String(rows[rowIndex][statusCol] || "").trim();

  if (currentStatus === "CLOSED") {
    throw new Error(
      "Pending delivery already closed"
    );
  }

  sheet
    .getRange(rowIndex + 1, statusCol + 1)
    .setValue("CLOSED");

  sheet
    .getRange(rowIndex + 1, closedAtCol + 1)
    .setValue(new Date());

  sheet
    .getRange(rowIndex + 1, closedByCol + 1)
    .setValue(by || "");

  return {
    success: true,
    pendingId
  };
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    if (!e || !e.parameter || !e.parameter.action) {
      return json({ success: false, message: "Missing action" });
    }

    const action = String(e.parameter.action).trim();

    // 🔓 PUBLIC ONLY
    if (action === "products") {
      return json({
        success: true,
        data: getProducts()
      });
    }

    // ❌ admin ห้ามผ่าน GET
    return json({
      success: false,
      message: "Forbidden"
    });

  } catch (err) {
    return json({
      success: false,
      error: err.message
    });
  }
}



function getProducts() {
  // 🔒 FIX: Web App ต้องอ้างอิง Spreadsheet แบบชัดเจน
  const SPREADSHEET_ID = "1xeNVv2yLADoxuZQwYBEZNvlZnt9CKvJ40RLTwNOrQfU";
  const sheet = SpreadsheetApp
    .openById(SPREADSHEET_ID)
    .getSheetByName("Products");

  if (!sheet) {
    return [];
  }

  const rows = sheet.getDataRange().getValues();

  // ❌ ไม่มีข้อมูลจริง (มีแต่ header)
  if (rows.length < 2) {
    return [];
  }

  rows.shift(); // ลบ header

  return rows
    .map(r => {
      const productId = String(r[0] || "").trim();
      if (!productId) return null; // ❌ กันแถวว่าง

      const name   = String(r[1] || "").trim();
      const price  = Number(r[2]) || 0;
      const stock  = Number(r[3]) || 0;

      // 🖼 image (column E)
      const image =
        typeof r[4] === "string" && r[4].startsWith("http")
          ? r[4].trim()
          : "";

      // 🔓 active (column F)
      const active =
        r[5] === true ||
        r[5] === "TRUE" ||
        r[5] === 1 ||
        r[5] === "1";

      // 🟢 status (column G)
      const status = String(r[6] || "").trim();
      // 📝 note (column J)
      const note = String(r[9] || "").trim();

      return {
        productId,
        name,
        price,
        stock,
        image,
        active,
        status,
        note,
        detailsText: String(r[10] || "").trim(),     // column K
        compareImages: String(r[11] || "").trim()    // column L
      };
    })
    .filter(Boolean); // ✅ ตัด null (แถวพัง / ว่าง)
}

function getAdminProducts() {
  const sheet = getSS().getSheetByName("Products");
  if (!sheet) {
    return [];
  }

  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) {
    return [];
  }

  const headers = rows.shift().map(h => String(h || "").trim());

  // 🔴 CHANGED: อ่านตามชื่อ header ก่อน เพื่อกันคอลัมน์เลื่อน
  const col = {
    productId: headers.indexOf("productId"),
    name: headers.indexOf("name"),
    price: headers.indexOf("price"),
    stock: headers.indexOf("stock"),
    image: headers.indexOf("image"),
    active: headers.indexOf("active"),
    status: headers.indexOf("status"),
    note: headers.indexOf("note"),
    detailsText: headers.indexOf("detailsText"),
    compareImages: headers.indexOf("compareImages"),
    costPrice: headers.indexOf("costPrice")
  };

  return rows
    .map(r => {
      // 🔴 CHANGED: fallback เป็น index เดิม หาก header ยังไม่ตรง
      const productId = String(
        col.productId > -1 ? r[col.productId] : r[0]
      ).trim();
      if (!productId) return null;

      const rawImage = col.image > -1 ? r[col.image] : r[4];
      const rawActive = col.active > -1 ? r[col.active] : r[5];
      const rawCostPrice = col.costPrice > -1 ? r[col.costPrice] : r[12];

      return {
        productId,
        name: String(col.name > -1 ? r[col.name] : r[1] || "").trim(),
        price: Number(col.price > -1 ? r[col.price] : r[2]) || 0,
        stock: Number(col.stock > -1 ? r[col.stock] : r[3]) || 0,
        image:
          typeof rawImage === "string" && rawImage.startsWith("http")
            ? rawImage.trim()
            : "",
        active:
          rawActive === true ||
          rawActive === "TRUE" ||
          rawActive === 1 ||
          rawActive === "1",
        status: String(col.status > -1 ? r[col.status] : r[6] || "").trim(),
        note: String(col.note > -1 ? r[col.note] : r[9] || "").trim(),
        detailsText: String(
          col.detailsText > -1 ? r[col.detailsText] : r[10] || ""
        ).trim(),
        compareImages: String(
          col.compareImages > -1 ? r[col.compareImages] : r[11] || ""
        ).trim(),
        costPrice:
          rawCostPrice !== "" &&
          rawCostPrice !== null &&
          rawCostPrice !== undefined
            ? Number(rawCostPrice)
            : null
      };
    })
    .filter(Boolean);
}


function doPost(e) {
  try {
    /* ================= BASIC GUARD ================= */
    if (!e || !e.parameter) {
      throw new Error("Missing request data");
    }

    const params = e.parameter;
    const action = String(params.action || "").trim();

    Logger.log(
      "ACTION: " + action +
      " | TIMESTAMP: " + new Date().toISOString()
    );

    if (!action) {
      throw new Error("Missing action");
    }

    /* =================================================
       🔓 PUBLIC ACTION (NO AUTH REQUIRED)
    ================================================= */

    if (action === "adminLogin") {
      return json(
        adminLogin(
          params.username,
          params.password
        )
      );
    }

    if (action === "createOrder") {
      enforceCreateOrderRateLimit();
      enforceCreateOrderBodySize(e);

      const result = createOrder({
        items: JSON.parse(params.items || "[]"),
        poNumber: String(params.poNumber || "").trim()
      });

      return json({
        success: true,
        data: result
      });
    }

    /* =================================================
       🔐 AUTH REQUIRED (ALL BELOW NEED TOKEN)
    ================================================= */

    const token = String(params.token || "").trim();
    if (!token) {
      throw new Error("Missing token");
    }

    const auth = requireAuth(token); // ❗ invalid → throw

    /* ========= ADD PRODUCT ========= */
    if (action === "addProduct") {
      const result = addProduct(e, auth);
      return json({ success: true, data: result });
    }

    /* ========= ADMIN READ ========= */
    if (action === "orders") {
      return json({
        success: true,
        data: getOrders()
      });
    }

    if (action === "adminProducts") {
      return json({
        success: true,
        data: getAdminProducts()
      });
    }
    

    if (action === "stockLogs") {
      return json({
        success: true,
        data: getStockLogs()
      });
    }

    if (action === "pendingDeliveries") {
      return json({
        success: true,
        data: getPendingDeliveries()
      });
    }

    /* ========= ADMIN WRITE ========= */

    if (action === "approveOrder") {
      const result = approveOrder(
        token,
        params.orderId
      );
      
      return json({ success: true, data: result });
    }

    if (action === "rejectOrder") {
      const result = rejectOrder(token, params.orderId);
      return json({ success: true, data: result });
    }

    if (action === "stockIn") {
      const result = stockIn(token, {
        productId: params.productId,
        qty: Number(params.qty),
        reason: params.reason || ""
      });
      return json({ success: true, data: result });
    }

    if (action === "stockAdjust") {
      const result = stockAdjust(token, {
        productId: params.productId,
        newQty: Number(params.newQty),
        reason: params.reason || ""
      });
      return json({ success: true, data: result });
    }

    if (action === "createPendingDelivery") {
      const result = createPendingDelivery(
        {
          orderId: params.orderId,
          poNumber: params.poNumber,
          productId: params.productId,
          orderedQty: Number(params.orderedQty),
          qty: Number(params.qty),
          note: params.note || ""
        },
        auth.username
      );

      return json({
        success: true,
        data: result
      });
    }

    if (action === "closePendingDelivery") {
      const result = closePendingDelivery(
        params.pendingId,
        auth.username
      );

      return json({
        success: true,
        data: result
      });
    }    

    if (action === "updateProduct") {
      const result = updateProduct(e, auth);
      return json({ success: true, data: result });
    }

    if (action === "deleteProduct") {
      const result = deleteProduct(e, auth);
      return json({ success: true, data: result });
    }

    if (action === "uploadProductImage") {
      const result = uploadProductImage(e, auth);
      return json({ success: true, data: result });
    }

    if (action === "adminLogout") {
      const result = adminLogout(params.token);
      return json({ success: true, data: result });
    }

    if (action === "changePassword") {
      const result = changePassword(
        params.token,
        params.currentPassword,
        params.newPassword
      );
      return json({ success: true, data: result });
    }

    if (action === "generateOrderPDF") {
       const result = generateOrderPDF(token, params.orderId);
       return json({ success: true, data: result });
     }


    /* ========= MAINTENANCE ========= */
    if (action === "cleanupImages") {
      const dryRun =
        String(params.dryRun || "true") === "true";

      const result = cleanupUnusedImages({
        by: auth.username,
        dryRun
      });

      return json({
        success: true,
        data: result
      });
    }


    /* =================================================
       ❌ UNKNOWN ACTION
    ================================================= */
    throw new Error("Invalid action: " + action);

  } catch (err) {
    Logger.log("doPost ERROR: " + err.message);

    return json({
      success: false,
      error: err.message || "Server error"
    });
  }
}





function uploadProductImage(e, auth) {
  /* ================= AUTH GUARD ================= */
  if (!auth || !auth.username) {
    throw new Error("Unauthorized");
  }

  const by = auth.username;

  /* ================= PARAM GUARD ================= */
  if (!e || !e.parameter) {
    throw new Error("Invalid request (no parameters)");
  }

  /* ================= READ DATA ================= */
  let data = e.parameter.data;
  if (!data || typeof data !== "string") {
    throw new Error("No image data received");
  }

  const filename =
    String(e.parameter.filename || `product-${Date.now()}.jpg`).trim();

  const mimeType =
    String(e.parameter.mimeType || "image/jpeg").trim();

  /* ================= CLEAN BASE64 ================= */
  if (data.includes("base64,")) {
    data = data.split("base64,")[1];
  }

  /* ================= SIZE LIMIT GUARD ================= */
  const MAX_BASE64_SIZE = 5 * 1024 * 1024;
  if (data.length > MAX_BASE64_SIZE) {
    throw new Error("Image too large (max 5MB)");
  }

  /* ================= DECODE ================= */
  const bytes = Utilities.base64Decode(data);
  if (!bytes || bytes.length === 0) {
    throw new Error("Base64 decode failed");
  }

  const blob = Utilities.newBlob(bytes, mimeType, filename);

  /* ================= DRIVE ================= */
  const FOLDER_ID = "1un_A6DFFnknmEjx7LgACRKT7l8AgBBdK";
  const folder = DriveApp.getFolderById(FOLDER_ID);

  // 🔴 CHANGED: เก็บ reference เพื่อ rollback ถ้า share ไม่ผ่าน
  let file = null;

  try {
    file = folder.createFile(blob);
    file.setName(filename);

    const fileId = file.getId();
    const imageUrl = `https://lh3.googleusercontent.com/d/${fileId}=w800`;

    Logger.log(
      `Image uploaded by ${by}: ${filename} (${fileId})`
    );

    return {
      imageUrl,
      uploadedBy: by
    };

  } catch (err) {
    // 🔴 CHANGED: ป้องกัน orphan file ถ้า create สำเร็จแต่ share พัง
    if (file) {
      try {
        file.setTrashed(true);
      } catch (_) {}
    }

    const message = String(err && err.message ? err.message : err);

    // 🔴 CHANGED: โยน error ให้ตรงสาเหตุจริงมากขึ้น
    if (
      message.includes("Access denied") ||
      message.includes("DriveApp")
    ) {
      throw new Error(
        "อัปโหลดรูปไม่สำเร็จ: ระบบไม่สามารถเข้าถึง Google Drive ได้"
      );
    }

    throw err;
  }
}


function stockAdjust(token, data) {
  // 🔐 AUTH (หัวใจ)
  const auth = requireAuth(token);
  const by = auth.username;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
  // 🔒 FIX: ใช้ Spreadsheet เดียวทั้งระบบ
  const ss = getSS();
  const productSheet = ss.getSheetByName("Products");
  const logSheet = ss.getSheetByName("stock_logs");

  /* ================= VALIDATE INPUT ================= */
  if (
    !data ||
    !data.productId ||
    !Number.isInteger(data.newQty) ||
    data.newQty < 0
  ) {
    throw new Error("Invalid stock adjust data");
  }

  /* ================= LOAD PRODUCTS ================= */
  const products = productSheet.getDataRange().getValues();
  products.shift(); // remove header

  const idx = products.findIndex(
    r => String(r[0]).trim() === data.productId
  );

  if (idx === -1) {
    throw new Error("Product not found");
  }

  const before = Number(products[idx][3]);
  const after  = Number(data.newQty);
  const diff   = after - before;

  /* ================= UPDATE STOCK ================= */
  productSheet
    .getRange(idx + 2, 4) // column D: stock
    .setValue(after);

  /* ================= LOG ADJUST ================= */
  try {
    logSheet.appendRow([
      "LOG-" + Date.now(), // logId
      data.productId,      // productId
      "ADJUST",            // type
      diff,                // qty (delta)
      before,              // before
      after,               // after
      by,                  // by (จาก token)
      "",                  // orderId
      data.reason || "",   // reason
      new Date()           // timestamp
    ]);
  } catch (err) {
    try {
      productSheet
        .getRange(idx + 2, 4)
        .setValue(before);
    } catch (rollbackErr) {
      throw new Error(
        String(err.message || err) +
        " | Stock rollback failed: " +
        String(rollbackErr.message || rollbackErr)
      );
    }

    throw err;
  }

  return {
    success: true,
    adjustedBy: by
  };
  } finally {
    lock.releaseLock();
  }
}




const CREATE_ORDER_MAX_ITEMS = 50;
const CREATE_ORDER_MAX_QTY_PER_PRODUCT = 999;
const CREATE_ORDER_MAX_BODY_BYTES = 100 * 1024;
const CREATE_ORDER_RATE_LIMIT_MAX_REQUESTS = 30;
const CREATE_ORDER_RATE_LIMIT_WINDOW_MS = 2 * 60 * 1000;
const CREATE_ORDER_RATE_LIMIT_KEY = "CREATE_ORDER_RATE_LIMIT";

function enforceCreateOrderBodySize(e) {
  if (!e || !e.postData) {
    return;
  }

  let bodySize = Number(e.postData.length);

  if (!Number.isFinite(bodySize)) {
    bodySize = Utilities
      .newBlob(String(e.postData.contents || ""))
      .getBytes()
      .length;
  }

  if (bodySize > CREATE_ORDER_MAX_BODY_BYTES) {
    throw new Error("Order request is too large");
  }
}

function enforceCreateOrderRateLimit() {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const properties = PropertiesService.getScriptProperties();
    const rawState = properties.getProperty(CREATE_ORDER_RATE_LIMIT_KEY);
    const now = Date.now();
    let state = null;

    if (rawState) {
      try {
        state = JSON.parse(rawState);
      } catch (err) {
        state = null;
      }
    }

    if (
      !state ||
      !Number.isFinite(Number(state.windowStartedAt)) ||
      !Number.isInteger(Number(state.count)) ||
      Number(state.count) < 0 ||
      now - Number(state.windowStartedAt) >= CREATE_ORDER_RATE_LIMIT_WINDOW_MS
    ) {
      state = {
        count: 0,
        windowStartedAt: now
      };
    }

    if (Number(state.count) >= CREATE_ORDER_RATE_LIMIT_MAX_REQUESTS) {
      throw new Error("Too many order requests. Please try again later");
    }

    state.count = Number(state.count) + 1;
    properties.setProperty(
      CREATE_ORDER_RATE_LIMIT_KEY,
      JSON.stringify(state)
    );
  } finally {
    lock.releaseLock();
  }
}

function createOrder(data) {
  // 🔒 FIX: Web App ต้องอ้างอิง Spreadsheet แบบชัดเจน
  const SPREADSHEET_ID = "1xeNVv2yLADoxuZQwYBEZNvlZnt9CKvJ40RLTwNOrQfU";
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("Orders");
  
  if (!sheet) {
    throw new Error("Orders sheet not found");
  }

  /* ================= VALIDATE ROOT ================= */
  if (!data || !Array.isArray(data.items)) {
    throw new Error("Invalid order items");
  }

  if (data.items.length === 0) {
    throw new Error("Order must contain at least 1 item");
  }

  if (data.items.length > CREATE_ORDER_MAX_ITEMS) {
    throw new Error(
      `Order cannot contain more than ${CREATE_ORDER_MAX_ITEMS} items`
    );
  }

  const poNumber = String(data.poNumber || "").trim();

  /* ================= MERGE CLIENT ITEMS ================= */
  const requestedQtyByProduct = new Map();

  data.items.forEach((item, index) => {
    const productId = String(item.productId || "").trim();
    const qty       = Number(item.qty);
    
    if (
      !productId ||
      !Number.isInteger(qty) ||
      qty <= 0 ||
      qty > CREATE_ORDER_MAX_QTY_PER_PRODUCT
    ) {
      throw new Error(
        `Invalid item at index ${index}`
      );
    }

    const mergedQty =
      (requestedQtyByProduct.get(productId) || 0) + qty;

    if (mergedQty > CREATE_ORDER_MAX_QTY_PER_PRODUCT) {
      throw new Error(
        `Quantity exceeds limit for ${productId}`
      );
    }

    requestedQtyByProduct.set(productId, mergedQty);
  });

  /* ================= LOAD AUTHORITATIVE PRODUCTS ================= */
  const productSheet = ss.getSheetByName("Products");
  if (!productSheet) {
    throw new Error("Products sheet not found");
  }

  const productRows = productSheet.getDataRange().getValues();
  productRows.shift(); // remove header

  const cleanItems = [];
  let total = 0;

  requestedQtyByProduct.forEach((qty, productId) => {
    const pIndex = productRows.findIndex(
      r => String(r[0]).trim() === productId
    );

    if (pIndex === -1) {
      throw new Error("Product not found: " + productId);
    }

    const productRow = productRows[pIndex];
    const name = String(productRow[1] || "").trim();
    const price = Number(productRow[2]);
    const currentStock = Number(productRow[3]);
    const rawActive = productRow[5];
    const active =
      rawActive === true ||
      rawActive === "TRUE" ||
      rawActive === 1 ||
      rawActive === "1";

    if (!name || !Number.isFinite(price) || price < 0) {
      throw new Error("Invalid product data: " + productId);
    }

    if (!active) {
      throw new Error("Product is not active: " + productId);
    }

    if (!Number.isFinite(currentStock) || currentStock < qty) {
      throw new Error(
        `Stock not enough for ${productId} (remain ${currentStock})`
      );
    }

    cleanItems.push({
      productId,
      name,
      qty,
      price
    });

    total += qty * price;
  });

  /* ================= CREATE ORDER ================= */
  const orderId   = "ORD-" + Date.now();
  const status    = "PENDING";
  const createdAt = new Date();

  sheet.appendRow([
    orderId,
    JSON.stringify(cleanItems),
    total,
    status,
    createdAt,
    "",          // approvedAt
    "",          // approvedBy
    poNumber     // poNumber
  ]);

  return {
    success: true,
    orderId,
    total,
    items: cleanItems
  };
}



function json(payload) {
  // 🔒 normalize null / undefined
  if (payload === null || payload === undefined) {
    payload = {};
  }

  // 🔴 ห้ามส่ง array ตรง ๆ (กันพังเงียบ)
  if (Array.isArray(payload)) {
    throw new Error("json() payload must be an object, not array");
  }

  // 🔒 enforce success flag (ค่าเดียว เชื่อถือได้)
  const success =
    payload.success === false || payload.error ? false : true;

  // ❌ ห้ามให้ success ซ้อน
  const { success: _ignored, ...rest } = payload;

  return ContentService
    .createTextOutput(
      JSON.stringify({
        success,
        ...rest
      })
    )
    .setMimeType(ContentService.MimeType.JSON);
}




function approveOrder(token, orderId) {
  /* ================= AUTH (GLOBAL) ================= */
  const auth = requireAuth(token);
  const by = auth.username;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    /* ================= SPREADSHEET ================= */
    const SPREADSHEET_ID = "1xeNVv2yLADoxuZQwYBEZNvlZnt9CKvJ40RLTwNOrQfU";
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const orderSheet   = ss.getSheetByName("Orders");
    const productSheet = ss.getSheetByName("Products");
    const logSheet     = ss.getSheetByName("stock_logs");

    /* ================= LOAD ORDER ================= */
    const orders = orderSheet.getDataRange().getValues();
    orders.shift();

    const idx = orders.findIndex(r => r[0] === orderId);
    if (idx === -1) {
      throw new Error("Order not found");
    }

    const rowNumber = idx + 2;
    const orderRow = orders[idx];
    const status = String(orderRow[3] || "").trim().toUpperCase();

    if (status === "APPROVED") {
      return {
        success: true,
        approvedBy: String(orderRow[6] || by)
      };
    }

    if (status === "REJECTED") {
      throw new Error("Order already rejected");
    }

    if (status !== "PENDING") {
      throw new Error("Invalid order status");
    }

    const items = JSON.parse(orderRow[1] || "[]");
    if (!Array.isArray(items) || items.length === 0) {
      throw new Error("Invalid order items");
    }

    /* ================= LOAD PRODUCTS ================= */
    const products = productSheet.getDataRange().getValues();
    products.shift();

    /* ================= MERGE + VALIDATE STOCK ================= */
    const requiredQtyByProduct = new Map();

    items.forEach((item, itemIndex) => {
      const productId = String(item.productId || "").trim();
      const qty = Number(item.qty);

      if (
        !productId ||
        !Number.isInteger(qty) ||
        qty <= 0
      ) {
        throw new Error("Invalid order item at index " + itemIndex);
      }

      requiredQtyByProduct.set(
        productId,
        (requiredQtyByProduct.get(productId) || 0) + qty
      );
    });

    const mutations = [];

    requiredQtyByProduct.forEach((qty, productId) => {
      const pIndex = products.findIndex(
        r => String(r[0]).trim() === productId
      );

      if (pIndex === -1) {
        throw new Error("Product not found: " + productId);
      }

      const currentStock = Number(products[pIndex][3]);
      if (!Number.isFinite(currentStock) || currentStock < qty) {
        throw new Error(
          `Stock not enough for ${productId} (remain ${currentStock})`
        );
      }

      mutations.push({
        productId,
        rowNumber: pIndex + 2,
        qty,
        before: currentStock,
        after: currentStock - qty
      });
    });

    /* ================= APPLY + COMPENSATE ON FAILURE ================= */
    const appliedMutations = [];
    const createdLogIds = [];

    try {
      mutations.forEach(mutation => {
        productSheet
          .getRange(mutation.rowNumber, 4)
          .setValue(mutation.after);
        appliedMutations.push(mutation);
      });

      const logTimestamp = new Date();

      mutations.forEach(mutation => {
        const logId = "LOG-" + Utilities.getUuid();
        createdLogIds.push(logId);

        logSheet.appendRow([
          logId,
          mutation.productId,
          "OUT",
          mutation.qty,
          mutation.before,
          mutation.after,
          by,
          orderId,
          "",
          logTimestamp
        ]);
      });

      orderSheet
        .getRange(rowNumber, 4, 1, 4)
        .setValues([[
          "APPROVED",
          orderRow[4],
          new Date(),
          by
        ]]);
    } catch (err) {
      let rollbackError = null;

      if (createdLogIds.length > 0) {
        try {
          const createdLogIdSet = new Set(createdLogIds);
          const lastLogRow = logSheet.getLastRow();

          if (lastLogRow >= 2) {
            const logIds = logSheet
              .getRange(2, 1, lastLogRow - 1, 1)
              .getValues();

            const rowsToDelete = [];

            logIds.forEach((row, index) => {
              if (createdLogIdSet.has(String(row[0]))) {
                rowsToDelete.push(index + 2);
              }
            });

            rowsToDelete
              .sort((a, b) => b - a)
              .forEach(logRowNumber => {
                logSheet.deleteRow(logRowNumber);
              });
          }
        } catch (logRollbackErr) {
          rollbackError = logRollbackErr;
        }
      }

      if (appliedMutations.length > 0) {
        try {
          appliedMutations.reverse().forEach(mutation => {
            productSheet
              .getRange(mutation.rowNumber, 4)
              .setValue(mutation.before);
          });
        } catch (stockRollbackErr) {
          rollbackError = rollbackError || stockRollbackErr;
        }
      }

      if (rollbackError) {
        throw new Error(
          String(err.message || err) +
          " | Rollback failed: " +
          String(rollbackError.message || rollbackError)
        );
      }

      throw err;
    }

    return {
      success: true,
      approvedBy: by
    };
  } finally {
    lock.releaseLock();
  }
}




function rejectOrder(token, orderId) {
  // 🔐 AUTH
  const auth = requireAuth(token);
  const by = auth.username;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
  // 🔒 FIX: ใช้ Spreadsheet เดียวทั้งระบบ
  const ss = getSS();
  const orderSheet = ss.getSheetByName("Orders");

  const rows = orderSheet.getDataRange().getValues();
  rows.shift(); // header

  const orderRowIndex = rows.findIndex(r => r[0] === orderId);
  if (orderRowIndex === -1) {
    throw new Error("Order not found");
  }

  const rowNumber = orderRowIndex + 2;
  const status = rows[orderRowIndex][3];

  // ❌ ป้องกัน reject ซ้ำ
  if (status !== "PENDING") {
    throw new Error("Order already processed");
  }

  // ===== Update status + audit =====
  orderSheet.getRange(rowNumber, 4).setValue("REJECTED");  // status
  orderSheet.getRange(rowNumber, 6).setValue(new Date()); // rejectedAt
  orderSheet.getRange(rowNumber, 7).setValue(by);         // rejectedBy

  return {
    success: true,
    rejectedBy: by
  };
  } finally {
    lock.releaseLock();
  }
}



 function getStockLogs() {
   // 🔒 FIX: ใช้ Spreadsheet เดียวทั้งระบบ
   const sheet = getSS().getSheetByName("stock_logs");

   if (!sheet) {
     return [];
   }

   const rows = sheet.getDataRange().getValues();
   if (rows.length < 2) return [];
  const headers = rows.shift().map(h => String(h || "").trim());

   return rows.map(r => {
     let obj = {};
     headers.forEach((h, i) => obj[h] = r[i]);

    // ✅ backward compatibility:
    // ชีตเก่าไม่มีคอลัมน์ reason และ IN/ADJUST เคยเอา reason ไปใส่ไว้ในช่อง orderId
    if (!("reason" in obj) || obj.reason === "" || obj.reason == null) {
      if (
        (obj.type === "IN" || obj.type === "ADJUST") &&
        obj.orderId &&
        !String(obj.orderId).startsWith("ORD-")
      ) {
        obj.reason = obj.orderId;
        obj.orderId = "";
      }
    }

     return obj;
   });
 }


function updateProduct(e, auth) {
  try {
    // 🔐 AUTH GUARD
    if (!auth || !auth.username) {
      throw new Error("Unauthorized");
    }

    const by = auth.username;

    /* ================= READ PARAM ================= */
    const oldProductId = String(e.parameter.oldProductId || "").trim();
    const newProductId = String(e.parameter.newProductId || "").trim();
    const name   = String(e.parameter.name || "").trim();
    const price  = Number(e.parameter.price);
    const rawCostPrice = String(e.parameter.costPrice ?? "").trim();
    const costPrice = rawCostPrice === "" ? null : Number(rawCostPrice);
    const stock  = Number(e.parameter.stock);
    const image  = String(e.parameter.image || "").trim();
    const status = String(e.parameter.status || "").trim();
    const note   = String(e.parameter.note || "").trim();
    const detailsText   = String(e.parameter.detailsText || "").trim();
    const compareImages = String(e.parameter.compareImages || "").trim();


    /* ================= VALIDATE ================= */
    if (!oldProductId || !newProductId || !name ||
       !Number.isInteger(price) ||
       costPrice === null ||
       !Number.isInteger(costPrice) ||
       !Number.isInteger(stock)) {
      throw new Error("Invalid product data");
    }

    if (price < 0 || costPrice < 0 || stock < 0) {
      throw new Error("Price, cost price and stock must be >= 0");
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);

    try {
      const ss = getSS();
      const sh = ss.getSheetByName("Products");
      if (!sh) {
        throw new Error("Sheet Products not found");
      }

      const data = sh.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === oldProductId) {

          // 🔒 อ่าน active และ stock ล่าสุดภายใน lock
          const originalRow = data[i].slice();
          const currentActive = data[i][5]; // column F
          const oldStock = Number(data[i][3]);
          const stockChanged = oldStock !== stock;
          const logSheet = stockChanged
            ? ss.getSheetByName("stock_logs")
            : null;

          if (stockChanged && !logSheet) {
            throw new Error("Sheet stock_logs not found");
          }

          try {
            // ===== HANDLE SKU CHANGE =====
            if (oldProductId !== newProductId) {

              // 🔍 Duplicate guard
              const exists = data.slice(1).some(
                r => String(r[0]).trim() === newProductId
              );
              if (exists) {
                throw new Error("SKU already exists");
              }

              // 🔄 Update SKU column A
              sh.getRange(i + 1, 1).setValue(newProductId);
            }

            /* ================= UPDATE ================= */
            sh.getRange(i + 1, 2).setValue(name);          // B: name
            sh.getRange(i + 1, 3).setValue(price);         // C: price
            sh.getRange(i + 1, 4).setValue(stock);         // D: stock
            sh.getRange(i + 1, 5).setValue(image);         // E: image
            sh.getRange(i + 1, 6).setValue(currentActive); // F: active (คงเดิม)
            sh.getRange(i + 1, 7).setValue(status);        // G: status
            sh.getRange(i + 1, 10).setValue(note);         // J: note
            sh.getRange(i + 1, 11).setValue(detailsText);   // K: detailsText
            sh.getRange(i + 1, 12).setValue(compareImages); // L: compareImages
            sh.getRange(i + 1, 13).setValue(costPrice);     // M: costPrice

            if (stockChanged) {
              logSheet.appendRow([
                "LOG-" + Date.now(),
                newProductId,
                "ADJUST",
                stock - oldStock,
                oldStock,
                stock,
                by,
                "",
                "",
                new Date()
              ]);
            }
          } catch (err) {
            try {
              sh
                .getRange(i + 1, 1, 1, originalRow.length)
                .setValues([originalRow]);
            } catch (rollbackErr) {
              throw new Error(
                String(err.message || err) +
                " | Product rollback failed: " +
                String(rollbackErr.message || rollbackErr)
              );
            }

            throw err;
          }

          Logger.log(`Product ${newProductId} updated by ${by}`);

          return {
            updatedBy: by
          };
        }
      }

      throw new Error("ไม่พบสินค้า");
    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    Logger.log("updateProduct error: " + err);
    throw err;
  }
}

function deleteProduct(e, auth) {
  // 🔐 AUTH GUARD
  if (!auth || !auth.username) {
    throw new Error("Unauthorized");
  }

  const by = auth.username;

  const productId = String(e.parameter.productId || "").trim();
  if (!productId) {
    throw new Error("Missing productId");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getSS();
    const sh = ss.getSheetByName("Products");
    if (!sh) {
      throw new Error("Sheet Products not found");
    }

    const rows = sh.getDataRange().getValues();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === productId) {

        // ❌ HARD DELETE: ลบแถวออกจากชีตจริง
        sh.deleteRow(i + 1);

        Logger.log(`Product ${productId} hard-deleted by ${by}`);

        return {
          success: true,
          deletedBy: by
        };
      }
    }

    throw new Error("Product not found");
  } finally {
    lock.releaseLock();
  }
}

function stockIn(token, data) {
  // 🔐 AUTH
  const auth = requireAuth(token);
  const by = auth.username;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
  // 🔒 FIX: ใช้ Spreadsheet เดียวทั้งระบบ
  const ss = getSS();
  const productSheet = ss.getSheetByName("Products");
  const logSheet = ss.getSheetByName("stock_logs");

  // 🔍 Validate input
  if (!data.productId || !Number.isInteger(data.qty) || data.qty <= 0) {
    throw new Error("Invalid stock in data");
  }

  const products = productSheet.getDataRange().getValues();
  products.shift(); // header

  const idx = products.findIndex(r => r[0] === data.productId);
  if (idx === -1) {
    throw new Error("Product not found");
  }

  const before = Number(products[idx][3]);
  const after = before + data.qty;

  // ===== Update stock =====
  productSheet
    .getRange(idx + 2, 4) // column stock
    .setValue(after);

  // ===== Log IN =====
  try {
    logSheet.appendRow([
      "LOG-" + Date.now(),   // logId
      data.productId,        // productId
      "IN",                  // type
      data.qty,              // qty
      before,                // before
      after,                 // after
      by,                    // by (จาก token)
      "",                    // orderId
      data.reason || "",     // reason
      new Date()             // timestamp
    ]);
  } catch (err) {
    try {
      productSheet
        .getRange(idx + 2, 4)
        .setValue(before);
    } catch (rollbackErr) {
      throw new Error(
        String(err.message || err) +
        " | Stock rollback failed: " +
        String(rollbackErr.message || rollbackErr)
      );
    }

    throw err;
  }

  return {
    success: true,
    by
  };
  } finally {
    lock.releaseLock();
  }
}



function adminLogin(username, password) {
  cleanupExpiredSessions();

  const ss = getSS();
  const adminSheet   = ss.getSheetByName("admins");
  const sessionSheet = ss.getSheetByName("sessions");

  if (!adminSheet || !sessionSheet) {
    return {
      success: false,
      message: "System not ready"
    };
  }

  // normalize input
  username = String(username || "").trim();
  password = String(password || "").trim();

  if (!username || !password) {
    return {
      success: false,
      message: "Username หรือ Password ไม่ถูกต้อง"
    };
  }

  if (isLoginRateLimited(username)) {
    return {
      success: false,
      message: "พยายามเข้าสู่ระบบมากเกินไป กรุณารอสักครู่"
    };
  }

  const lastRow = adminSheet.getLastRow();
  if (lastRow < 2) {
    return {
      success: false,
      message: "No admin configured"
    };
  }

  // ===== HASH PASSWORD (UPGRADE) =====
  const passwordHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password
  )
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2))
    .join('');

  // load admins (username, password_hash)
  const admins = adminSheet
    .getRange(2, 1, lastRow - 1, 2)
    .getValues()
    .map(r => [
      String(r[0]).trim(), // username
      String(r[1]).trim()  // password_hash
    ]);

  const found = admins.find(
    r => r[0] === username && r[1] === passwordHash
  );

  if (!found) {
    recordFailedLoginAttempt(username);

    return {
      success: false,
      message: "Username หรือ Password ไม่ถูกต้อง"
    };
  }

  clearLoginAttempts(username);

  // ===== remove old session of this user =====
  const sessionRows = sessionSheet.getDataRange().getValues();
  const keep = [sessionRows[0]]; // header

  for (let i = 1; i < sessionRows.length; i++) {
    if (sessionRows[i][1] !== username) {
      keep.push(sessionRows[i]);
    }
  }

  sessionSheet.clearContents();
  sessionSheet
    .getRange(1, 1, keep.length, keep[0].length)
    .setValues(keep);

  // ===== create new session =====
  const token = Utilities.getUuid();
  const expiredAt = new Date(Date.now() + 12 * 60 * 60 * 1000); // 12 ชม.

  sessionSheet.appendRow([
    token,
    username,
    expiredAt
  ]);

  return {
    success: true,
    token,
    username,
    expiredAt
  };
}

function adminLogout(token) {
  token = String(token || "").trim();

  if (!token) {
    throw new Error("Missing token");
  }

  const sessionSheet = getSS().getSheetByName("sessions");

  if (!sessionSheet) {
    throw new Error("Session system not ready");
  }

  const lastRow = sessionSheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("Invalid token");
  }

  const tokens = sessionSheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues();

  const sessionIndex = tokens.findIndex(
    row => String(row[0]).trim() === token
  );

  if (sessionIndex === -1) {
    throw new Error("Invalid token");
  }

  sessionSheet.deleteRow(sessionIndex + 2);

  return {
    success: true
  };
}

function changePassword(token, currentPassword, newPassword) {
  const auth = requireAuth(token);
  const username = auth.username;

  currentPassword = String(currentPassword || "").trim();
  newPassword = String(newPassword || "").trim();

  if (!currentPassword || !newPassword) {
    throw new Error("Missing password data");
  }

  if (newPassword.length < 4) {
    throw new Error("New password too short");
  }

  const ss = getSS();
  const adminSheet = ss.getSheetByName("admins");
  const sessionSheet = ss.getSheetByName("sessions");

  if (!adminSheet || !sessionSheet) {
    throw new Error("System not ready");
  }

  const currentHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    currentPassword
  )
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2))
    .join('');

  const newHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    newPassword
  )
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2))
    .join('');

  const lastRow = adminSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error("No admin configured");
  }

  const rows = adminSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const rowIndex = rows.findIndex(r =>
    String(r[0]).trim() === username &&
    String(r[1]).trim() === currentHash
  );

  if (rowIndex === -1) {
    throw new Error("Current password is incorrect");
  }

  adminSheet.getRange(rowIndex + 2, 2).setValue(newHash);

  const sessionRows = sessionSheet.getDataRange().getValues();
  const keep = [sessionRows[0]];

  for (let i = 1; i < sessionRows.length; i++) {
    if (String(sessionRows[i][1]).trim() !== username) {
      keep.push(sessionRows[i]);
    }
  }

  sessionSheet.clearContents();
  sessionSheet
    .getRange(1, 1, keep.length, keep[0].length)
    .setValues(keep);

  return {
    success: true,
    username
  };
}

function requireAuth(token) {
  if (!token) {
    throw new Error("Missing token");
  }

  const ss = getSS();
  const sheet = ss.getSheetByName("sessions");

  if (!sheet) {
    throw new Error("Session system not ready");
  }

  const lastRow = sheet.getLastRow();

  // 🔒 GUARD: ไม่มี session จริง (มีแต่ header หรือว่าง)
  if (lastRow < 2) {
    throw new Error("Invalid token");
  }

  const rows = sheet
    .getRange(2, 1, lastRow - 1, 3)
    .getValues();

  const now = new Date();

  for (const [savedToken, username, expiredAt] of rows) {
    if (String(savedToken) === String(token)) {

      if (!expiredAt || new Date(expiredAt) <= now) {
        throw new Error("Session expired");
      }

      return {
        username,
        token,
        role: "admin", // 🔒 เตรียมไว้
        expiredAt: new Date(expiredAt)
      };
    }
  }

  throw new Error("Invalid token");
}




function cleanupExpiredSessions() {
  // 🔒 FIX: ใช้ Spreadsheet เดียวทั้งระบบ
  const ss = getSS();
  const sheet = ss.getSheetByName("sessions");

  // ===== GUARD =====
  if (!sheet) {
    Logger.log("cleanupExpiredSessions: sessions sheet not found");
    return;
  }

  const rows = sheet.getDataRange().getValues();

  // ถ้าไม่มีข้อมูล หรือมีแค่ header
  if (rows.length <= 1) {
    Logger.log("cleanupExpiredSessions: no session rows to clean");
    return;
  }

  const now = new Date();
  const keep = [rows[0]]; // header

  let removed = 0;

  for (let i = 1; i < rows.length; i++) {
    const expiredAt = rows[i][2];

    // guard ค่า expiredAt
    if (!expiredAt) {
      removed++;
      continue;
    }

    if (new Date(expiredAt) > now) {
      keep.push(rows[i]);
    } else {
      removed++;
    }
  }

  // ===== WRITE BACK =====
  sheet.clearContents();
  sheet
    .getRange(1, 1, keep.length, keep[0].length)
    .setValues(keep);

  Logger.log(
    `cleanupExpiredSessions: removed ${removed} expired session(s)`
  );
}

// ================= LOGIN RATE LIMIT =================
const LOGIN_RATE_LIMIT_MAX_ATTEMPTS = 5;
const LOGIN_RATE_LIMIT_WINDOW_MS = 15 * 60 * 1000;

function getLoginAttemptKey(username) {
  const normalizedUsername =
    String(username || "").trim().toLowerCase();

  const hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    normalizedUsername
  )
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2))
    .join('');

  return "LOGIN_ATTEMPTS_" + hash;
}

function isLoginRateLimited(username) {
  const key = getLoginAttemptKey(username);
  const properties = PropertiesService.getScriptProperties();
  const rawState = properties.getProperty(key);

  if (!rawState) {
    return false;
  }

  let state;

  try {
    state = JSON.parse(rawState);
  } catch (err) {
    properties.deleteProperty(key);
    return false;
  }

  const firstFailedAt = Number(state.firstFailedAt);
  const count = Number(state.count);

  if (
    !Number.isFinite(firstFailedAt) ||
    !Number.isInteger(count) ||
    Date.now() - firstFailedAt >= LOGIN_RATE_LIMIT_WINDOW_MS
  ) {
    properties.deleteProperty(key);
    return false;
  }

  return count >= LOGIN_RATE_LIMIT_MAX_ATTEMPTS;
}

function recordFailedLoginAttempt(username) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const key = getLoginAttemptKey(username);
    const properties = PropertiesService.getScriptProperties();
    const rawState = properties.getProperty(key);
    const now = Date.now();
    let state = null;

    if (rawState) {
      try {
        state = JSON.parse(rawState);
      } catch (err) {
        state = null;
      }
    }

    if (
      !state ||
      !Number.isFinite(Number(state.firstFailedAt)) ||
      !Number.isInteger(Number(state.count)) ||
      Number(state.count) < 0 ||
      now - Number(state.firstFailedAt) >= LOGIN_RATE_LIMIT_WINDOW_MS
    ) {
      state = {
        count: 0,
        firstFailedAt: now
      };
    }

    state.count = Number(state.count) + 1;
    properties.setProperty(key, JSON.stringify(state));
  } finally {
    lock.releaseLock();
  }
}

function clearLoginAttempts(username) {
  PropertiesService
    .getScriptProperties()
    .deleteProperty(getLoginAttemptKey(username));
}


function getOrders() {
  // 🔒 FIX: ใช้ Spreadsheet เดียวกับทั้งระบบ
  const sheet = getSS().getSheetByName("Orders");

  if (!sheet) {
    return [];
  }

  const rows = sheet.getDataRange().getValues();

  // มีแต่ header
  if (rows.length < 2) {
    return [];
  }

  const headers = rows.shift();

  return rows.map(r => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = r[i];
    });

    // 🔧 parse items จาก string → array
    if (typeof obj.items === "string") {
      try {
        obj.items = JSON.parse(obj.items);
      } catch (e) {
        obj.items = [];
      }
    }

    return obj;
  });
}


function addProduct(e, auth) {
  if (!auth || !auth.username) {
    throw new Error("Unauthorized");
  }

  const by = auth.username;

  /* ================= READ PARAM ================= */
  const id     = String(e.parameter.productId || "").trim();
  const name   = String(e.parameter.name || "").trim();
  const price  = Number(e.parameter.price);
  const rawCostPrice = String(e.parameter.costPrice ?? "").trim();
  const costPrice = rawCostPrice === "" ? null : Number(rawCostPrice);
  const stock  = Number(e.parameter.stock);
  const active = true; // สินค้าใหม่เปิดขายเสมอ
  const image  = String(e.parameter.image || "").trim();
  let   status = String(e.parameter.status || "ready").trim();
  const note   = String(e.parameter.note || "").trim();

  /* ================= VALIDATE ================= */
  if (!id || !name ||
     !Number.isInteger(price) ||
     costPrice === null ||
     !Number.isInteger(costPrice) ||
     !Number.isInteger(stock)) {
    throw new Error("Invalid product data");
  }

  if (price < 0 || costPrice < 0 || stock < 0) {
    throw new Error("Price, cost price and stock must be >= 0");
  }

  // 🔒 auto status: stock = 0 → out
  if (stock === 0) {
    status = "out";
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    // 🔒 ใช้ Spreadsheet เดียวกับทั้งระบบ
    const ss = getSS();

    const sh = ss.getSheetByName("Products");
    if (!sh) throw new Error("Products sheet not found");

    // 🆕 LOG SHEET
    const logSheet = ss.getSheetByName("stock_logs");
    if (!logSheet) throw new Error("stock_logs sheet not found");

    // 🔍 กัน SKU ซ้ำภายใน lock
    const rows = sh.getDataRange().getValues();
    const exists = rows.slice(1).some(r => String(r[0]).trim() === id);
    if (exists) {
      throw new Error("SKU already exists");
    }

    const createdRowNumber = sh.getLastRow() + 1;

    /* ================= ADD PRODUCT ================= */
    sh.appendRow([
      id,          // A productId
      name,        // B name
      price,       // C price
      stock,       // D stock
      image,       // E image
      active,      // F active
      status,      // G status  ✅ เพิ่มใหม่
      new Date(),  // H createdAt
      by,          // I createdBy
      note,        // J note
      "",          // K detailsText
      "",          // L compareImages
      costPrice    // M costPrice
    ]);

    try {
      /* ================= LOG CREATE ================= */
      logSheet.appendRow([
        "LOG-" + Date.now(), // logId
        id,                  // productId
        "CREATE",            // type
        stock,               // qty
        0,                   // before
        stock,               // after
        by,                  // by
        "",                  // orderId
        "CREATE_PRODUCT",    // reason
        new Date()           // timestamp
      ]);
    } catch (err) {
      try {
        const createdProductId = String(
          sh.getRange(createdRowNumber, 1).getValue()
        ).trim();

        if (createdProductId !== id) {
          throw new Error("Created product row no longer matches");
        }

        sh.deleteRow(createdRowNumber);
      } catch (rollbackErr) {
        throw new Error(
          String(err.message || err) +
          " | Product create rollback failed: " +
          String(rollbackErr.message || rollbackErr)
        );
      }

      throw err;
    }

    return {
      success: true,
      createdBy: by
    };
  } finally {
    lock.releaseLock();
  }
}

function cleanupUnusedImages({ by, dryRun }) {
  const ss = getSS();
  const productSheet = ss.getSheetByName("Products");
  if (!productSheet) {
    throw new Error("Products sheet not found");
  }

  /* ================= COLLECT USED IMAGE IDS ================= */
  const rows = productSheet.getDataRange().getValues();
  rows.shift(); // header

  const usedFileIds = new Set();

  rows.forEach(r => {
    const imageUrl = String(r[4] || "").trim();
    if (!imageUrl) return;

    // extract fileId from lh3 url
    const match = imageUrl.match(/\/d\/([^=]+)/);
    if (match && match[1]) {
      usedFileIds.add(match[1]);
    }
  });

  /* ================= SCAN DRIVE FOLDER ================= */
  const FOLDER_ID = "1un_A6DFFnknmEjx7LgACRKT7l8AgBBdK";
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  let scanned = 0;
  let orphan = [];
  let removed = 0;

  while (files.hasNext()) {
    const file = files.next();
    scanned++;

    const fileId = file.getId();
    if (!usedFileIds.has(fileId)) {
      orphan.push({
        fileId,
        name: file.getName()
      });

      if (!dryRun) {
        file.setTrashed(true);
        removed++;
      }
    }
  }

  /* ================= LOG ================= */
  Logger.log(
    `[GC] by=${by} dryRun=${dryRun} scanned=${scanned} orphan=${orphan.length}`
  );

  return {
    dryRun,
    scanned,
    orphanCount: orphan.length,
    removed,
    orphanFiles: orphan
  };
}

function generateOrderPDF(token, orderId) {
  const auth = requireAuth(token);
  const by = auth.username;

  if (!orderId) {
    throw new Error("Missing orderId");
  }

  const ss = getSS();
  const sheet = ss.getSheetByName("Orders");
  if (!sheet) {
    throw new Error("Orders sheet not found");
  }

  const rows = sheet.getDataRange().getValues();
  rows.shift(); // remove header

  const idx = rows.findIndex(r => r[0] === orderId);
  if (idx === -1) {
    throw new Error("Order not found");
  }

  const row = rows[idx];

  let items = [];
  try {
    items = JSON.parse(row[1] || "[]");
    if (!Array.isArray(items)) items = [];
  } catch (err) {
    items = [];
  }
  const total = Number(row[2]) || 0;
  const status = String(row[3] || "");
  const createdAtRaw = row[4] ? new Date(row[4]) : new Date();
  const createdAt = Utilities.formatDate(
    createdAtRaw,
    "Asia/Bangkok",
    "dd/MM/yyyy HH:mm:ss"
  );

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  const html = `
    <html>
      <head>
        <meta charset="UTF-8">
        <style>
          body { font-family: Arial; padding: 20px; }
          table { width: 100%; border-collapse: collapse; margin-top: 20px; }
          th, td { border: 1px solid #ddd; padding: 6px; font-size: 12px; }
          th { background: #f2f2f2; }
        </style>
      </head>
      <body>
        <h2>Order ${escapeHtml(orderId)}</h2>
        <p>Status: ${escapeHtml(status)}</p>
        <p>Date: ${createdAt}</p>

        <table>
          <tr>
            <th>Product</th>
            <th>Qty</th>
            <th>Price</th>
            <th>Total</th>
          </tr>
          ${items.map(i => `
            <tr>
              <td>${escapeHtml(i.name)}</td>
              <td>${i.qty}</td>
              <td>${i.price}</td>
              <td>${Number(i.qty) * Number(i.price)}</td>
            </tr>
          `).join("")}
        </table>

        <h3>Grand Total: ${total}</h3>
        <p>Generated by: ${escapeHtml(by)}</p>
      </body>
    </html>
  `;

  const blob = Utilities.newBlob(html, "text/html")
    .getAs("application/pdf")
    .setName(`Order-${String(orderId).replace(/[^a-zA-Z0-9-_]/g, "")}.pdf`);

  const file = DriveApp.createFile(blob);

  file.setSharing(
    DriveApp.Access.ANYONE,
    DriveApp.Permission.VIEW
  );

  return {
    url: file.getDownloadUrl(),
    fileId: file.getId()
  };
}

