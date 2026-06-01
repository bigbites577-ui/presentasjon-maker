const form = document.querySelector("#briefForm");
const statusEl = document.querySelector("#status");
const imageGrid = document.querySelector("#imageGrid");
const slideList = document.querySelector("#slideList");
const downloadBtn = document.querySelector("#downloadBtn");

let selectedImage = "";
let lastPayload = null;

function getPayload() {
  return {
    topic: document.querySelector("#topic").value.trim(),
    criteria: document.querySelector("#criteria").value.trim(),
    style: document.querySelector("#style").value,
    slideCount: Number(document.querySelector("#slideCount").value || 7),
    notesLevel: document.querySelector("#notesLevel").value,
    emoji: document.querySelector("#emoji").checked,
    selectedImage,
  };
}

function setStatus(text) {
  statusEl.textContent = text;
}

function renderSlides(slides) {
  slideList.innerHTML = "";
  slides.forEach((slide, index) => {
    const notePreview = slide.notes
      ? slide.notes.replace(/\s+/g, " ").slice(0, 180)
      : "Speaker notes will be generated here.";
    const li = document.createElement("li");
    li.innerHTML = `
      <span>${String(index + 1).padStart(2, "0")}</span>
      <div>
        <b>${slide.title}</b>
        <p>${slide.bullets.slice(0, 2).join(" · ")}</p>
        <p class="note-preview">${notePreview}${slide.notes && slide.notes.length > 180 ? "..." : ""}</p>
      </div>
    `;
    slideList.appendChild(li);
  });
}

function renderImages(images) {
  imageGrid.classList.remove("empty");
  imageGrid.innerHTML = "";
  images.forEach((image, index) => {
    const card = document.createElement("button");
    card.type = "button";
    card.className = `visual-card${index === 0 ? " selected" : ""}`;
    card.innerHTML = `<img alt="${image.label}" src="${image.dataUri}"><p>${image.label}</p>`;
    card.addEventListener("click", () => {
      selectedImage = image.dataUri;
      document.querySelectorAll(".visual-card").forEach(item => item.classList.remove("selected"));
      card.classList.add("selected");
      downloadBtn.disabled = false;
    });
    imageGrid.appendChild(card);
  });
  selectedImage = images[0]?.dataUri || "";
}

async function postJson(url, payload) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    const error = await response.json().catch(() => ({ error: "Something went wrong" }));
    throw new Error(error.error || "Something went wrong");
  }
  return response;
}

form.addEventListener("submit", async event => {
  event.preventDefault();
  selectedImage = "";
  const payload = getPayload();
  if (!payload.topic) {
    setStatus("Need topic");
    return;
  }

  setStatus("Generating");
  downloadBtn.disabled = true;
  try {
    const response = await postJson("/api/images", payload);
    const data = await response.json();
    renderImages(data.images);
    renderSlides(data.slides);
    lastPayload = { ...payload, selectedImage };
    downloadBtn.disabled = false;
    setStatus("Ready");
  } catch (error) {
    setStatus("Error");
    imageGrid.classList.add("empty");
    imageGrid.innerHTML = `<div class="empty-state">${error.message}</div>`;
  }
});

downloadBtn.addEventListener("click", async () => {
  const payload = { ...(lastPayload || getPayload()), ...getPayload(), selectedImage };
  if (!payload.topic) {
    setStatus("Need topic");
    return;
  }

  setStatus("Building");
  downloadBtn.disabled = true;
  try {
    const response = await postJson("/api/pptx", payload);
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const safeName = payload.topic.replace(/[^a-z0-9æøå]+/gi, "-").replace(/^-|-$/g, "") || "presentation";
    link.href = url;
    link.download = `${safeName}.pptx`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setStatus("Downloaded");
  } catch (error) {
    setStatus("Error");
    alert(error.message);
  } finally {
    downloadBtn.disabled = false;
  }
});

document.querySelector("#topic").value = "The future of sustainable transport in cities";
document.querySelector("#criteria").value = "Audience: classmates. Include problem, current solutions, benefits, challenges, and a clear ending. Tone: confident and easy to understand.";
