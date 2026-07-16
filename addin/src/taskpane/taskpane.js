/* global Office, PowerPoint */

const CATALOG_URL = "https://yazevstas.github.io/Presentation-Library/catalog.json";
let slidesCatalog = [];

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    loadCatalog();
    document.getElementById("searchInput").addEventListener("input", handleSearch);
  }
});

async function loadCatalog() {
  try {
    const response = await fetch(CATALOG_URL);
    slidesCatalog = await response.json();
    
    document.getElementById("loading").style.display = "none";
    document.getElementById("catalog").style.display = "flex";
    
    renderSlides(slidesCatalog);
  } catch (error) {
    document.getElementById("loading").innerText = "Ошибка загрузки каталога. Проверьте сеть.";
    console.error("Ошибка загрузки каталога:", error);
  }
}

function renderSlides(slides) {
  const container = document.getElementById("catalog");
  container.innerHTML = ""; 

  if (slides.length === 0) {
    container.innerHTML = "<div class='loading'>Слайды не найдены</div>";
    return;
  }

  slides.forEach(slide => {
    const card = document.createElement("div");
    card.className = "slide-card";

    card.innerHTML = `
      <img src="${slide.preview_url}" class="slide-preview" alt="${slide.title}">
      <div class="slide-info">
        <p class="slide-title">${slide.title}</p>
        <p class="slide-category">${slide.category}</p>
        <button class="insert-btn" id="btn-${slide.id}">Вставить слайд</button>
      </div>
    `;

    container.appendChild(card);

    document.getElementById(`btn-${slide.id}`).addEventListener("click", () => {
      // Передаем также ID кнопки, чтобы менять её цвет и текст
      insertSlide(slide.slide_url, `btn-${slide.id}`);
    });
  });
}

function handleSearch(event) {
  const query = event.target.value.toLowerCase();
  const filteredSlides = slidesCatalog.filter(slide => 
    slide.title.toLowerCase().includes(query) || 
    slide.category.toLowerCase().includes(query)
  );
  renderSlides(filteredSlides);
}

// УЛУЧШЕННАЯ ФУНКЦИЯ ВСТАВКИ
async function insertSlide(slideTxtUrl, buttonId) {
  const btn = document.getElementById(buttonId);
  const originalText = btn.innerText;
  
  // Меняем дизайн кнопки на время загрузки
  btn.innerText = "Загрузка слайда...";
  btn.style.backgroundColor = "#ff9e64"; // Оранжевый цвет

  try {
    // 1. Скачиваем текст Base64
    const response = await fetch(slideTxtUrl);
    if (!response.ok) throw new Error("Не удалось скачать файл слайда");
    let base64String = await response.text();

    // 2. ЖЕСТКАЯ ОЧИСТКА: удаляем все переносы строк, пробелы и невидимые символы
    base64String = base64String.replace(/[\r\n\s]+/g, "");
    // На всякий случай удаляем префикс data:image/..., если конвертер его добавил
    base64String = base64String.replace(/^data:.*?;base64,/, "");

    // 3. Вставка через Office.js
    await PowerPoint.run(async (context) => {
      context.presentation.insertSlidesFromBase64(base64String, {
        targetSlideId: null, 
        action: "KeepSourceFormatting" 
      });
      await context.sync();
    });

    // 4. Если успех — кнопка зеленая
    btn.innerText = "Успешно вставлено!";
    btn.style.backgroundColor = "#9ece6a"; 
    
    // Возвращаем кнопку в исходное состояние через 3 секунды
    setTimeout(() => {
      btn.innerText = originalText;
      btn.style.backgroundColor = "#00d2ff";
    }, 3000);

  } catch (error) {
    console.error("Подробная ошибка:", error);
    
    // 5. Если ошибка — кнопка красная, никаких alert()
    btn.innerText = "Ошибка вставки";
    btn.style.backgroundColor = "#f7768e"; 
    
    setTimeout(() => {
      btn.innerText = originalText;
      btn.style.backgroundColor = "#00d2ff";
    }, 3000);
  }
}

// Функция фильтрации по боковому меню
function filterCategory(categoryName) {
  // Подсвечиваем активный пункт меню
  const items = document.querySelectorAll('.menu-item');
  items.forEach(item => item.classList.remove('active'));
  event.currentTarget.classList.add('active');

  // Фильтруем каталог
  if (categoryName === '') {
    renderSlides(slidesCatalog); // Показать всё
  } else {
    const filtered = slidesCatalog.filter(slide => slide.category.includes(categoryName));
    renderSlides(filtered);
  }
}