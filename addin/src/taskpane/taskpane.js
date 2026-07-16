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
    document.getElementById("catalog").style.display = "grid"; // Используем grid для 2 колонок
    
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
    // Если ничего не найдено, растягиваем сообщение на 2 колонки
    container.innerHTML = "<div class='loading' style='grid-column: span 2;'>Слайды не найдены</div>";
    return;
  }

  slides.forEach(slide => {
    const card = document.createElement("div");
    card.className = "slide-card";

    // Здесь та самая обновленная кнопка с плюсиком (+)
    card.innerHTML = `
      <img src="${slide.preview_url}" class="slide-preview" alt="${slide.title}">
      <div class="slide-info">
        <div class="text-block">
          <p class="slide-title">${slide.title}</p>
          <p class="slide-category">${slide.category}</p>
        </div>
        <button class="insert-btn" id="btn-${slide.id}" title="Вставить слайд">+</button>
      </div>
    `;

    container.appendChild(card);

    document.getElementById(`btn-${slide.id}`).addEventListener("click", () => {
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

// Улучшенная функция вставки с анимацией для маленькой кнопки
async function insertSlide(slideTxtUrl, buttonId) {
  const btn = document.getElementById(buttonId);
  
  // Меняем плюсик на песочные часы (Загрузка)
  btn.innerText = "⏳";
  btn.style.backgroundColor = "#ff9e64"; // Оранжевый
  btn.style.color = "#ffffff";

  try {
    const response = await fetch(slideTxtUrl);
    if (!response.ok) throw new Error("Не удалось скачать файл");
    let base64String = await response.text();

    base64String = base64String.replace(/[\r\n\s]+/g, "");
    base64String = base64String.replace(/^data:.*?;base64,/, "");

    await PowerPoint.run(async (context) => {
      context.presentation.insertSlidesFromBase64(base64String, {
        targetSlideId: null, 
        action: "KeepSourceFormatting" 
      });
      await context.sync();
    });

    // Меняем на галочку (Успех)
    btn.innerText = "✓";
    btn.style.backgroundColor = "#9ece6a"; // Зеленый
    
    // Возвращаем плюсик через 3 секунды
    setTimeout(() => {
      btn.innerText = "+";
      btn.style.backgroundColor = ""; // Сброс стиля
      btn.style.color = "";
    }, 3000);

  } catch (error) {
    console.error("Подробная ошибка:", error);
    
    // Меняем на крестик (Ошибка)
    btn.innerText = "✖";
    btn.style.backgroundColor = "#f7768e"; // Красный
    
    setTimeout(() => {
      btn.innerText = "+";
      btn.style.backgroundColor = "";
      btn.style.color = "";
    }, 3000);
  }
}

// Функция фильтрации по боковому меню
function filterCategory(categoryName) {
  // Выделяем активный пункт меню синим
  const e = window.event;
  if (e && e.currentTarget) {
      const items = document.querySelectorAll('.menu-item');
      items.forEach(item => item.classList.remove('active'));
      e.currentTarget.classList.add('active');
  }

  // Фильтруем каталог
  if (categoryName === '') {
    renderSlides(slidesCatalog); // Показать всё
  } else {
    const filtered = slidesCatalog.filter(slide => slide.category.includes(categoryName));
    renderSlides(filtered);
  }
}

// Функция сворачивания/разворачивания бокового меню
function toggleSidebar() {
  const sidebar = document.getElementById('sidebar');
  sidebar.classList.toggle('collapsed');
}