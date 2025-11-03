const crm = "crm@amp-aero.com"; // 👈 Change this to your CRM Bcc address

function getBccAsync(item) {
  return new Promise((resolve, reject) => {
    item.bcc.getAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        resolve(res.value || []);
      } else {
        reject(res.error);
      }
    });
  });
}

function addBccAsync(item, address) {
  return new Promise((resolve, reject) => {
    item.bcc.addAsync([{ emailAddress: address }], (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(res.error);
      }
    });
  });
}

export async function addCrmBcc() {
  const item = Office.context.mailbox.item;
  const currentBcc = await getBccAsync(item);
  const exists = currentBcc.some(
    (a) => (a.emailAddress || a).toLowerCase() === crm.toLowerCase()
  );
  if (!exists) {
    await addBccAsync(item, crm);
  }
}

// Register the button action
Office.actions.associate("addCrmBcc", async (event) => {
  try {
    await addCrmBcc();
  } catch (err) {
    console.error("Failed to add CRM Bcc:", err);
  } finally {
    event.completed();
  }
});
