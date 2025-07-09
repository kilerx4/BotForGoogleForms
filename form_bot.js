const xlsx = require('xlsx');
const puppeteer = require('puppeteer');
const path = require('path');
const stringSimilarity = require('string-similarity');

// === ПАРАМЕТРЫ ИЗ КОМАНДНОЙ СТРОКИ ===
const args = process.argv.slice(2);
const START_ROW = args[0] ? parseInt(args[0], 10) : 0;
const MAX_ROWS = args[1] ? (args[1] === 'null' || args[1] === '' ? null : parseInt(args[1], 10)) : null;
const HEADLESS = args[2] ? (args[2] === 'true' || args[2] === '1') : false;

// === НАСТРОЙКИ ===
const FORM_URL = 'YOUR_GOOGLE_FORM_URL_HERE'; // Замените на вашу ссылку Google формы
const EXCEL_FILE = path.join(__dirname, 'moscow_sport_data.xlsx');
const DELAY_BETWEEN = [2000, 4000]; // Задержка между отправками (мс)

// Функция для расчета схожести строк с поддержкой синонимов
function calculateSimilarity(str1, str2) {
  const s1 = str1.toLowerCase();
  const s2 = str2.toLowerCase();
  
  // Проверяем точное совпадение
  if (s1 === s2) return 1.0;
  
  // Проверяем включение
  if (s1.includes(s2) || s2.includes(s1)) return 0.8;
  
  // Проверяем синонимы
  const synonyms = {
    'редко': ['реже одного раза в месяц', 'раз в месяц'],
    'смартфон': ['телефон'],
    'актуальная': ['всегда актуальна'],
    'нет': ['не доверяю', 'никогда'],
    'никакой логики': ['другое'],
    'периодически вылетает': ['других особенностей'],
    'часто зависает': ['других особенностей'],
    'если появится геймификация': ['система бонусов'],
    'whatsapp': ['социальные сети'],
    'telegram': ['социальные сети']
  };
  
  for (const [synonym, variants] of Object.entries(synonyms)) {
    if (s1.includes(synonym) || s2.includes(synonym)) {
      for (const variant of variants) {
        if (s1.includes(variant) || s2.includes(variant)) {
          return 0.9;
        }
      }
    }
  }
  
  // Обычное сравнение по символам
  let matches = 0;
  for (let i = 0; i < Math.min(s1.length, s2.length); i++) {
    if (s1[i] === s2[i]) matches++;
  }
  return matches / Math.max(s1.length, s2.length);
}

function sleep(ms) {
  return new Promise(res => setTimeout(res, ms));
}

function randomDelay(min, max) {
  return sleep(Math.floor(Math.random() * (max - min + 1)) + min);
}

async function getFormQuestions(page) {
  console.log('🔍 Ищу вопросы в форме...');
  
  // Попробуем разные селекторы для поиска вопросов
  const selectors = [
    'div.freebirdFormviewerComponentsQuestionBaseRoot',
    'div[data-item-id]',
    'div[role="listitem"]',
    'div.freebirdFormviewerComponentsQuestionBaseTitle',
    'div[data-params*="question"]'
  ];
  
  for (const selector of selectors) {
    const questions = await page.evaluate((sel) => {
      const elements = document.querySelectorAll(sel);
      console.log(`Найдено элементов с селектором ${sel}:`, elements.length);
      return Array.from(elements).map(el => ({
        text: el.innerText.trim(),
        html: el.innerHTML.substring(0, 200) + '...',
        selector: sel,
        element: el
      }));
    }, selector);
    
    if (questions.length > 0) {
      console.log(`✅ Найдено ${questions.length} вопросов с селектором: ${selector}`);
      
      // Извлекаем только заголовки вопросов (до первого переноса строки или звездочки)
      const questionHeaders = questions.map(q => {
        const lines = q.text.split('\n');
        const header = lines[0].trim();
        return {
          header: header,
          fullText: q.text,
          element: q.element
        };
      });
      
      questionHeaders.forEach((q, i) => {
        console.log(`  ${i + 1}. "${q.header}"`);
      });
      
      return questionHeaders;
    }
  }
  
  // Если ничего не нашли, выведем всю структуру страницы
  console.log('❌ Вопросы не найдены. Анализирую структуру страницы...');
  const pageStructure = await page.evaluate(() => {
    const allDivs = Array.from(document.querySelectorAll('div')).slice(0, 50);
    return allDivs.map(div => ({
      className: div.className,
      id: div.id,
      text: div.innerText.substring(0, 100),
      role: div.getAttribute('role')
    }));
  });
  
  console.log('Структура страницы (первые 50 div):');
  pageStructure.forEach((item, i) => {
    console.log(`  ${i + 1}. class="${item.className}" id="${item.id}" role="${item.role}" text="${item.text}"`);
  });
  
  return [];
}

async function fillForm(page, row, columns, formQuestions) {
  console.log(`📝 Заполняю форму для строки с данными...`);
  
  for (const col of columns) {
    const value = row[col];
    // Пропускаем служебную колонку "Отметка времени"
    if (col.toLowerCase().includes('отметка времени')) {
      console.log(`⏭️ Пропускаю: ${col}`);
      continue;
    }
    
    // Для вопроса о функциях в мобильном приложении не пропускаем пустые значения
    const isMobileFunctionsQuestion = col.toLowerCase().includes('функций') && col.toLowerCase().includes('мобильном приложении');
    
    if (!value && !isMobileFunctionsQuestion) {
      console.log(`⏭️ Пропускаю: ${col} (пустое значение)`);
      continue;
    }
    
    const excelQuestion = col;
    console.log(`\n🔍 Ищу вопрос: "${excelQuestion}"`);
    
    // Для вопроса о функциях в мобильном приложении, если значение пустое или содержит только запятую, заполняем случайным текстом
    let finalValue = value;
    if (isMobileFunctionsQuestion && (!value || value.trim() === '' || value.trim() === ',')) {
      const randomFunctions = [
        'Карта мероприятий, Запись на события',
        'Уведомления о новых событиях, Поиск по местоположению',
        'Календарь тренировок, Статистика активности',
        'Чат с тренерами, Видео-уроки',
        'Интеграция с фитнес-трекерами, Социальные функции',
        'Онлайн-запись на секции, Планировщик маршрутов',
        'Персональные рекомендации, Система достижений'
      ];
      finalValue = randomFunctions[Math.floor(Math.random() * randomFunctions.length)];
      console.log(`📄 Пустое значение для функций, заполняю случайным: "${finalValue}"`);
    } else {
      console.log(`📄 Значение: "${value}"`);
    }

    if (formQuestions.length === 0) {
      console.log(`❌ Нет вопросов в форме для сопоставления`);
      continue;
    }

    // Fuzzy matching: ищем лучший вопрос в форме по заголовкам
    try {
      const questionHeaders = formQuestions.map(q => q.header);
      const { bestMatch } = stringSimilarity.findBestMatch(excelQuestion, questionHeaders);
      const questionHeader = bestMatch.target;
      const rating = bestMatch.rating;
      
      console.log(`🎯 Лучшее совпадение: "${questionHeader}" (рейтинг: ${rating.toFixed(2)})`);
      
      if (rating < 0.4) {
        console.log(`❗ Не найден похожий вопрос для: ${excelQuestion}`);
        continue;
      }
      
      // Находим соответствующий элемент формы
      const questionData = formQuestions.find(q => q.header === questionHeader);
      if (!questionData) {
        console.log(`❗ Не найден элемент для вопроса: ${questionHeader}`);
        continue;
      }

      console.log(`✅ Найден блок вопроса, заполняю...`);

      // Пробуем найти radio/checkbox в этом блоке
      const radioClicked = await page.evaluate((questionText, value) => {
        // Ищем родительский блок вопроса
        const blocks = Array.from(document.querySelectorAll('div[role="listitem"]'));
        const targetBlock = blocks.find(block => block.innerText.includes(questionText));
        
        if (!targetBlock) {
          console.log(`Блок вопроса не найден: ${questionText}`);
          return false;
        }
        
        // Ищем все radio кнопки и их текстовые метки
        const radios = targetBlock.querySelectorAll('[role=radio]');
        console.log(`Найдено ${radios.length} radio кнопок`);
        
        for (let i = 0; i < radios.length; i++) {
          const radio = radios[i];
          
          // Ищем текст варианта ответа - он может быть в соседнем элементе
          let optionText = '';
          
          // Попробуем найти текст в родительском элементе
          const parent = radio.closest('div[role="radio"]') || radio.parentElement;
          if (parent) {
            optionText = parent.innerText.trim();
          }
          
          // Если не нашли, попробуем найти в соседних элементах
          if (!optionText) {
            const siblings = Array.from(radio.parentElement.children);
            for (const sibling of siblings) {
              if (sibling !== radio && sibling.innerText.trim()) {
                optionText = sibling.innerText.trim();
                break;
              }
            }
          }
          
          // Если все еще нет текста, попробуем найти по aria-label
          if (!optionText && radio.getAttribute('aria-label')) {
            optionText = radio.getAttribute('aria-label');
          }
          
          console.log(`Проверяю radio ${i + 1}: "${optionText}"`);
          
          if (optionText && optionText.includes(value)) {
            console.log(`Найдено совпадение! Кликаю: "${optionText}"`);
            radio.click();
            return true;
          }
        }
        
        // Если точное совпадение не найдено, попробуем fuzzy matching
        console.log(`Точное совпадение не найдено, пробую fuzzy matching...`);
        const allOptions = [];
        for (let i = 0; i < radios.length; i++) {
          const radio = radios[i];
          let optionText = '';
          
          const parent = radio.closest('div[role="radio"]') || radio.parentElement;
          if (parent) {
            optionText = parent.innerText.trim();
          }
          
          if (!optionText) {
            const siblings = Array.from(radio.parentElement.children);
            for (const sibling of siblings) {
              if (sibling !== radio && sibling.innerText.trim()) {
                optionText = sibling.innerText.trim();
                break;
              }
            }
          }
          
          if (optionText) {
            allOptions.push({ text: optionText, radio: radio });
          }
        }
        
        // Специальная обработка для вопроса о поиске
        if (questionText.includes('Как часто вы пользуетесь поиском на платформе')) {
          console.log('🔍 Специальная обработка для вопроса о поиске...');
          console.log(`📄 Ищем ответ: "${value}"`);
          
          // Собираем варианты ответов так же, как в основном коде
          const searchOptions = [];
          for (let i = 0; i < radios.length; i++) {
            const radio = radios[i];
            let optionText = '';
            
            // Ищем текст в родительском элементе
            const parent = radio.closest('div[role="radio"]') || radio.parentElement;
            if (parent) {
              optionText = parent.innerText.trim();
            }
            
            // Если не нашли, ищем в соседних элементах
            if (!optionText) {
              const siblings = Array.from(radio.parentElement.children);
              for (const sibling of siblings) {
                if (sibling !== radio && sibling.innerText.trim()) {
                  optionText = sibling.innerText.trim();
                  break;
                }
              }
            }
            
            // Если все еще не нашли, используем aria-label
            if (!optionText && radio.getAttribute('aria-label')) {
              optionText = radio.getAttribute('aria-label');
            }
            
            if (optionText) {
              searchOptions.push({ text: optionText, radio: radio });
              console.log(`📋 Вариант ${i + 1}: "${optionText}"`);
            }
          }
          
          console.log(`📋 Всего вариантов: ${searchOptions.length}`);
          
          // Сопоставление для ответов о поиске
          const searchMapping = {
            'редко': 'Иногда',
            'часто': 'Всегда', 
            'всегда': 'Всегда',
            'иногда': 'Иногда',
            'никогда': 'Никогда',
            'постоянно': 'Всегда',
            'периодически': 'Иногда'
          };
          
          const normalizedValue = value.toLowerCase().trim();
          const mappedOption = searchMapping[normalizedValue];
          
          console.log(`🎯 Нормализованное значение: "${normalizedValue}"`);
          console.log(`🎯 Сопоставленный вариант: "${mappedOption}"`);
          
          if (mappedOption) {
            const targetOption = searchOptions.find(opt => opt.text === mappedOption);
            if (targetOption) {
              console.log(`✅ Специальное сопоставление: "${value}" → "${mappedOption}"`);
              targetOption.radio.click();
              return true;
            } else {
              console.log(`❌ Вариант "${mappedOption}" не найден в доступных опциях`);
              console.log(`📋 Доступные варианты: ${searchOptions.map(opt => opt.text).join(', ')}`);
            }
          } else {
            console.log(`❌ Нет сопоставления для значения "${normalizedValue}"`);
          }
        }
        
        // Обычный fuzzy matching для остальных случаев
        if (allOptions.length > 0) {
          let bestMatch = null;
          let bestScore = 0;
          
          for (const option of allOptions) {
            const score = calculateSimilarity(value.toLowerCase(), option.text.toLowerCase());
            console.log(`Fuzzy match: "${value}" → "${option.text}" (рейтинг: ${score.toFixed(2)})`);
            
            if (score > bestScore && score >= 0.7) {
              bestScore = score;
              bestMatch = option;
            }
          }
          
          if (bestMatch) {
            console.log(`✅ Fuzzy match: "${value}" → "${bestMatch.text}" (рейтинг: ${bestScore.toFixed(2)})`);
            bestMatch.radio.click();
            return true;
          }
        }
        
        // Если fuzzy matching не сработал, попробуем принудительное сопоставление для обязательных вопросов
        console.log('🔍 Принудительное сопоставление для обязательного вопроса...');
        
                  // Сопоставления для разных типов вопросов
          const forcedMappings = {
            // Частота использования платформы
            'как часто вы пользуетесь платформой': {
              'редко': ['Раз в месяц', 'Реже одного раза в месяц'],
              'часто': ['Ежедневно', 'Несколько раз в неделю'],
              'всегда': ['Ежедневно', 'Несколько раз в неделю'],
              'иногда': ['Несколько раз в неделю', 'Раз в месяц'],
              'никогда': ['Впервые', 'Реже одного раза в месяца'],
              'постоянно': ['Ежедневно', 'Несколько раз в неделю'],
              'периодически': ['Несколько раз в неделю', 'Раз в месяц']
            },
          // Устройство
          'с какого устройства': {
            'смартфон': 'Телефон',
            'телефон': 'Телефон',
            'компьютер': 'Компьютер/ноутбук',
            'ноутбук': 'Компьютер/ноутбук',
            'планшет': 'Планшет',
            'другое': 'Другое'
          },
          // Трудности с меню
          'какие трудности': {
            'никакой логики': 'Другое',
            'никакой логики в расположении': 'Другое',
            'путаюсь': 'Другое',
            'все запутано': 'Другое',
            'сложно найти': 'Нужную информацию сложно найти',
            'много пунктов': 'Слишком много пунктов',
            'непонятно': 'Непонятны названия разделов'
          },
          // Мотивация
          'что могло бы мотивировать': {
            'геймификация': 'Система бонусов',
            'если появится геймификация': 'Система бонусов',
            'если станет проще': 'Упрощение процесса',
            'бонусы': 'Система бонусов',
            'рекомендации': 'Персональные рекомендации'
          },
          // Актуальность
          'как вы оцениваете актуальность': {
            'актуальная': 'Всегда актуальна',
            'часто устаревшая': 'Иногда устаревшая',
            'не соответствует': 'Часто не соответствует действительности',
            'не хватает деталей': 'Иногда не хватает деталей'
          },
          // Каналы уведомлений
          'какие каналы для уведомлений': {
            'whatsapp': 'Социальные сети',
            'telegram': 'Социальные сети',
            'whatsapp, telegram, социальные сети': 'Социальные сети',
            'push': 'Push-уведомления',
            'email': 'Email',
            'sms': 'SMS',
            'смс': 'SMS'
          },
          // Доверие
          'доверяете ли вы платформе': {
            'нет': 'Не доверяю',
            'да': 'Полностью',
            'скорее нет': 'Не доверяю',
            'частично': 'Частично'
          },
          // Трудности
          'испытывали ли вы трудности': {
            'периодически вылетает': 'Других особенностей',
            'часто зависает': 'Других особенностей',
            'периодически вылетает, часто зависает': 'Других особенностей',
            'интерфейс неудобный': 'Других особенностей',
            'неудобный интерфейс': 'Других особенностей'
          }
        };
        
        // Ищем подходящее сопоставление
        for (const [questionPattern, mappings] of Object.entries(forcedMappings)) {
          if (questionText.toLowerCase().includes(questionPattern)) {
            console.log(`🎯 Найдено сопоставление для: "${questionPattern}"`);
            
            const normalizedValue = value.toLowerCase().trim();
            const mappedOptions = mappings[normalizedValue];
            
            if (mappedOptions) {
              // Если это массив вариантов, выбираем случайный
              const mappedOption = Array.isArray(mappedOptions) 
                ? mappedOptions[Math.floor(Math.random() * mappedOptions.length)]
                : mappedOptions;
              
              // Ищем точное совпадение
              let targetOption = allOptions.find(opt => opt.text === mappedOption);
              
              // Если точное совпадение не найдено, ищем частичное
              if (!targetOption) {
                targetOption = allOptions.find(opt => 
                  opt.text.toLowerCase().includes(mappedOption.toLowerCase()) ||
                  mappedOption.toLowerCase().includes(opt.text.toLowerCase())
                );
              }
              
              if (targetOption) {
                console.log(`✅ Принудительное сопоставление: "${value}" → "${targetOption.text}"`);
                targetOption.radio.click();
                return true;
                              } else {
                  console.log(`❌ Вариант "${mappedOption}" не найден в доступных опциях`);
                  console.log(`📋 Доступные опции: [${allOptions.map(opt => `"${opt.text}"`).join(', ')}]`);
                }
            } else {
              console.log(`❌ Нет сопоставления для "${normalizedValue}" в "${questionPattern}"`);
            }
            break;
          }
        }
        
        // Если ничего не подошло, выбираем случайный вариант (чтобы форма отправилась)
        if (allOptions.length > 0) {
          // Добавляем больше рандома - исключаем крайние варианты для более реалистичных ответов
          let randomIdx;
          if (allOptions.length > 2) {
            // Для вопросов с 3+ вариантами исключаем первый и последний (часто это крайности)
            randomIdx = 1 + Math.floor(Math.random() * (allOptions.length - 2));
          } else {
            // Для вопросов с 1-2 вариантами выбираем любой
            randomIdx = Math.floor(Math.random() * allOptions.length);
          }
          console.log(`🎲 Выбираю случайный вариант #${randomIdx + 1}: "${allOptions[randomIdx].text}"`);
          allOptions[randomIdx].radio.click();
          return true;
        }
        
        // Если это обязательный вопрос (есть звездочка), принудительно выбираем случайный вариант
        if (questionText.includes('*')) {
          console.log(`🎲 Принудительный выбор для обязательного вопроса с звездочкой`);
          const radios = targetBlock.querySelectorAll('[role="radio"]');
          console.log(`🎲 Найдено ${radios.length} radio кнопок для принудительного выбора`);
          if (radios.length > 0) {
            // Добавляем больше рандома
            let randomIdx;
            if (radios.length > 2) {
              // Для вопросов с 3+ вариантами используем взвешенный рандом
              if (Math.random() < 0.7) {
                // 70% вероятность выбрать средние варианты
                randomIdx = 1 + Math.floor(Math.random() * (radios.length - 2));
              } else {
                // 30% вероятность выбрать любые варианты
                randomIdx = Math.floor(Math.random() * radios.length);
              }
            } else {
              randomIdx = Math.floor(Math.random() * radios.length);
            }
            console.log(`🎲 Принудительно кликаю случайную radio кнопку #${randomIdx + 1}`);
            radios[randomIdx].click();
            return true;
          }
        }
        
        return false;
      }, questionHeader, finalValue);
      
      if (radioClicked) {
        console.log(`✅ Заполнено radio: "${value}"`);
        continue;
      }

      // Пробуем найти input/textarea
      const inputFilled = await page.evaluate((questionText, value) => {
        const blocks = Array.from(document.querySelectorAll('div[role="listitem"]'));
        const targetBlock = blocks.find(block => block.innerText.includes(questionText));
        
        if (!targetBlock) {
          console.log(`❌ Блок для текстового поля не найден: "${questionText}"`);
          return false;
        }
        
        // Ищем разные типы текстовых полей
        let input = targetBlock.querySelector('input[type="text"], textarea, input[type="email"], input[type="url"]');
        
        // Если не нашли, ищем по более широким селекторам
        if (!input) {
          input = targetBlock.querySelector('input, textarea');
        }
        
        // Если все еще не нашли, ищем по роли
        if (!input) {
          input = targetBlock.querySelector('[role="textbox"], [contenteditable="true"]');
        }
        
        if (input) {
          console.log(`✅ Найден input/textarea для: "${questionText}"`);
          console.log(`📝 Заполняю текстом: "${value}"`);
          
          try {
            // Очищаем поле
            input.focus();
            input.click();
            
            // Очищаем содержимое
            input.value = '';
            input.textContent = '';
            
            // Заполняем текст
            const textToFill = String(value);
            
            // Способ 1: через value
            if (input.tagName === 'INPUT' || input.tagName === 'TEXTAREA') {
              input.value = textToFill;
              input.dispatchEvent(new Event('input', { bubbles: true }));
              input.dispatchEvent(new Event('change', { bubbles: true }));
            }
            
            // Способ 2: через textContent для contenteditable
            if (input.getAttribute('contenteditable') === 'true') {
              input.textContent = textToFill;
              input.dispatchEvent(new Event('input', { bubbles: true }));
            }
            
            // Способ 3: симуляция ввода символов
            input.focus();
            for (const char of textToFill) {
              input.value += char;
              input.dispatchEvent(new Event('input', { bubbles: true }));
            }
            
            console.log(`✅ Текстовое поле заполнено успешно`);
            return true;
          } catch (error) {
            console.log(`❌ Ошибка при заполнении текстового поля: ${error.message}`);
            return false;
          }
        } else {
          console.log(`❌ Текстовое поле не найдено в блоке для: "${questionText}"`);
          console.log(`🔍 Доступные элементы в блоке:`, targetBlock.innerHTML.substring(0, 200) + '...');
          return false;
        }
      }, questionHeader, finalValue);
      
      if (inputFilled) {
        console.log(`✅ Заполнено input/textarea: "${value}"`);
        continue;
      }

      // Специальная обработка для вопросов с текстовыми ответами
      if (questionHeader.toLowerCase().includes('функций') && questionHeader.toLowerCase().includes('мобильном приложении') ||
          questionHeader.toLowerCase().includes('изменения') && questionHeader.toLowerCase().includes('платформу')) {
        console.log(`🔧 Специальная обработка для текстового вопроса: "${questionHeader}"`);
        
        const specialInputFilled = await page.evaluate((questionText, value) => {
          const blocks = Array.from(document.querySelectorAll('div[role="listitem"]'));
          const targetBlock = blocks.find(block => block.innerText.includes(questionText));
          
          if (!targetBlock) {
            console.log(`❌ Блок для специального вопроса не найден`);
            return false;
          }
          
          // Ищем все возможные текстовые поля
          const inputs = targetBlock.querySelectorAll('input, textarea, [role="textbox"], [contenteditable="true"]');
          console.log(`🔍 Найдено ${inputs.length} возможных текстовых полей`);
          
          for (const input of inputs) {
            try {
              console.log(`🔧 Пробую заполнить поле типа: ${input.tagName}`);
              
              // Фокусируемся на поле
              input.focus();
              input.click();
              
              // Очищаем поле
              if (input.tagName === 'INPUT' || input.tagName === 'TEXTAREA') {
                input.value = '';
              } else {
                input.textContent = '';
              }
              
              // Заполняем текст
              const textToFill = String(value);
              
              if (input.tagName === 'INPUT' || input.tagName === 'TEXTAREA') {
                input.value = textToFill;
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
              } else {
                input.textContent = textToFill;
                input.dispatchEvent(new Event('input', { bubbles: true }));
              }
              
              console.log(`✅ Специальное поле заполнено: "${textToFill}"`);
              return true;
            } catch (error) {
              console.log(`❌ Ошибка при заполнении специального поля: ${error.message}`);
              continue;
            }
          }
          
          return false;
        }, questionHeader, finalValue);
        
        if (specialInputFilled) {
          console.log(`✅ Специально заполнено текстовое поле: "${questionHeader}"`);
          continue;
        }
      }

      // Пробуем найти рейтинг (звезды)
      const ratingClicked = await page.evaluate((questionText, value) => {
        const blocks = Array.from(document.querySelectorAll('div[role="listitem"]'));
        const targetBlock = blocks.find(block => block.innerText.includes(questionText));
        
        if (!targetBlock) return false;
        
        const radios = targetBlock.querySelectorAll('[role=radio]');
        for (const radio of radios) {
          if (radio.innerText.trim() === String(value).trim()) {
            radio.click();
            return true;
          }
        }
        return false;
      }, questionHeader, finalValue);
      
      if (ratingClicked) {
        console.log(`✅ Заполнен рейтинг: "${value}"`);
        continue;
      }

      // Если ничего не нашли, но это обязательный вопрос - принудительно выбираем случайный вариант
      if (questionHeader.includes('*')) {
        console.log(`🚨 Принудительный выбор для обязательного вопроса: "${questionHeader}"`);
        
        // Попробуем найти и кликнуть любую radio кнопку в этом блоке
        const forcedClick = await page.evaluate((questionText) => {
          console.log(`🔍 Ищу принудительно radio кнопки для: "${questionText}"`);
          
          const blocks = Array.from(document.querySelectorAll('div[role="listitem"]'));
          console.log(`📋 Всего блоков на странице: ${blocks.length}`);
          
          // Пробуем разные способы поиска блока
          let targetBlock = null;
          
          // Способ 1: поиск по полному тексту
          for (const block of blocks) {
            if (block.innerText.includes(questionText)) {
              targetBlock = block;
              console.log(`✅ Найден блок по полному тексту`);
              break;
            }
          }
          
          // Способ 2: поиск по первой части вопроса
          if (!targetBlock) {
            const searchText = questionText.split('?')[0].substring(0, 20);
            console.log(`🔍 Ищу по части текста: "${searchText}"`);
            for (const block of blocks) {
              if (block.innerText.includes(searchText)) {
                targetBlock = block;
                console.log(`✅ Найден блок по части текста`);
                break;
              }
            }
          }
          
          // Способ 3: поиск по индексу (если это обязательный вопрос)
          if (!targetBlock) {
            console.log(`🔍 Ищу любой блок с radio кнопками и звездочкой`);
            for (const block of blocks) {
              if (block.innerText.includes('*')) {
                const radios = block.querySelectorAll('[role="radio"]');
                if (radios.length > 0) {
                  targetBlock = block;
                  console.log(`✅ Найден блок с radio кнопками и звездочкой`);
                  break;
                }
              }
            }
          }
          
          if (!targetBlock) {
            console.log(`❌ Блок вопроса не найден`);
            return false;
          }
          
          const radioButtons = targetBlock.querySelectorAll('[role="radio"]');
          console.log(`🚨 Найдено ${radioButtons.length} radio кнопок для принудительного выбора`);
          
          if (radioButtons.length > 0) {
            // Добавляем больше рандома
            let randomIdx;
            if (radioButtons.length > 2) {
              // Для вопросов с 3+ вариантами используем взвешенный рандом
              if (Math.random() < 0.65) {
                // 65% вероятность выбрать средние варианты
                randomIdx = 1 + Math.floor(Math.random() * (radioButtons.length - 2));
              } else {
                // 35% вероятность выбрать любые варианты
                randomIdx = Math.floor(Math.random() * radioButtons.length);
              }
            } else {
              randomIdx = Math.floor(Math.random() * radioButtons.length);
            }
            console.log(`🎲 Принудительно кликаю случайную radio кнопку #${randomIdx + 1}`);
            radioButtons[randomIdx].click();
            
            // Проверяем, что клик сработал
            setTimeout(() => {
              const isChecked = radioButtons[randomIdx].checked;
              console.log(`🔍 Проверка: radio кнопка выбрана = ${isChecked}`);
            }, 100);
            
            return true;
          }
          
          return false;
        }, questionHeader);
        
        if (forcedClick) {
          console.log(`✅ Принудительно заполнено radio для обязательного вопроса`);
          continue;
        } else {
          console.log(`❗ Не удалось заполнить обязательный вопрос: ${excelQuestion}`);
        }
      } else {
        console.log(`❗ Не удалось заполнить: ${excelQuestion}`);
      }
      
    } catch (error) {
      console.log(`❌ Ошибка при сопоставлении: ${error.message}`);
      continue;
    }
  }
  
  // Принудительное заполнение всех обязательных вопросов
  console.log('\n🚨 Принудительное заполнение всех обязательных вопросов...');
  
  const filledCount = await page.evaluate(() => {
    const blocks = Array.from(document.querySelectorAll('div[role="listitem"]'));
    let filled = 0;
    
    blocks.forEach((block, index) => {
      const text = block.innerText;
      if (text.includes('*')) {
        const radioButtons = block.querySelectorAll('[role="radio"]');
        const hasCheckedRadio = Array.from(radioButtons).some(radio => radio.getAttribute('aria-checked') === 'true');
        
        if (radioButtons.length > 0 && !hasCheckedRadio) {
          // Добавляем больше рандома в принудительное заполнение
          let randomIdx;
          if (radioButtons.length > 2) {
            // Для вопросов с 3+ вариантами используем взвешенный рандом
            // 60% вероятность выбрать средние варианты, 40% - любые
            if (Math.random() < 0.6) {
              randomIdx = 1 + Math.floor(Math.random() * (radioButtons.length - 2));
            } else {
              randomIdx = Math.floor(Math.random() * radioButtons.length);
            }
          } else {
            randomIdx = Math.floor(Math.random() * radioButtons.length);
          }
          console.log(`🎲 Принудительно заполняю вопрос ${index + 1} случайным вариантом #${randomIdx + 1}`);
          radioButtons[randomIdx].click();
          filled++;
        }
      }
    });
    
    return filled;
  });
  
  console.log(`✅ Принудительно заполнено ${filledCount} обязательных вопросов`);
  
  // Небольшая пауза после заполнения
  await sleep(1000);
}

async function main() {
  console.log('🚀 Запуск бота для заполнения Google Формы');
  console.log('=' .repeat(50));
  
  // Чтение Excel
  console.log('📊 Читаю Excel файл...');
  const workbook = xlsx.readFile(EXCEL_FILE);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(sheet);
  const columns = Object.keys(data[0]);
  const endRow = MAX_ROWS ? START_ROW + MAX_ROWS : data.length;

  console.log(`📈 Всего строк в Excel: ${data.length}`);
  console.log(`📋 Заголовки из Excel:`, columns);

  // Открываем форму и получаем все вопросы
  console.log('🌐 Открываю Google Форму...');
  const browser = await puppeteer.launch({ 
    headless: HEADLESS, 
    defaultViewport: null,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });
  const page = await browser.newPage();
  
  // Добавляем логирование консоли браузера
  page.on('console', msg => console.log('🌐 Браузер:', msg.text()));
  
  await page.goto(FORM_URL, { waitUntil: 'networkidle2' });
  console.log('✅ Форма загружена');
  await sleep(3000);
  
  const formQuestions = await getFormQuestions(page);
  console.log(`\n📝 Вопросы, найденные в форме (${formQuestions.length}):`, formQuestions);

  console.log(`\n⚙️ Настройки выполнения:`);
  console.log(`   - Будет обработано: с ${START_ROW + 1} по ${endRow}`);
  console.log(`   - Режим headless: ${HEADLESS}`);

  for (let i = START_ROW; i < endRow && i < data.length; i++) {
    const row = data[i];
    console.log(`\n${'='.repeat(50)}`);
    console.log(`📝 Заполняю запись #${i + 1}`);
    console.log(`${'='.repeat(50)}`);
    
    await page.goto(FORM_URL, { waitUntil: 'networkidle2' });
    await sleep(2000);
    await fillForm(page, row, columns, formQuestions);
    
    // Нажать кнопку "Отправить"
    console.log('📤 Ищу кнопку отправки...');
    const submitBtn = await page.evaluateHandle(() => {
      // Попробуем разные способы найти кнопку отправки
      const selectors = [
        'div[role="button"]:has(span:contains("Отправить"))',
        'div[aria-label="Submit"]',
        'div[jsname="M2UYVd"]',
        'div.uArJ5e:has(.NPEfkd)',
        'span.NPEfkd:contains("Отправить")',
        'div[role="button"] span:contains("Отправить")'
      ];
      
      for (const selector of selectors) {
        try {
          const elements = document.querySelectorAll(selector);
          console.log(`Найдено ${elements.length} элементов с селектором: ${selector}`);
          if (elements.length > 0) {
            return elements[0];
          }
        } catch (e) {
          console.log(`Ошибка с селектором ${selector}:`, e.message);
        }
      }
      
      // Если не нашли по селекторам, ищем по тексту
      const allButtons = Array.from(document.querySelectorAll('div[role="button"], span[role="button"]'));
      console.log(`Найдено ${allButtons.length} кнопок`);
      allButtons.forEach((btn, i) => {
        console.log(`Кнопка ${i + 1}: "${btn.innerText}"`);
      });
      
      const submitButton = allButtons.find(btn => 
        btn.innerText && btn.innerText.includes('Отправить')
      );
      
      return submitButton || null;
    });
    
    const isSubmitDefined = await page.evaluate(btn => !!btn, submitBtn);
    if (isSubmitDefined) {
      console.log('✅ Кнопка отправки найдена, кликаю...');
      
      // Попробуем несколько способов клика
      try {
        // Способ 1: Обычный клик
        await submitBtn.click();
        console.log('✅ Клик выполнен');
        
        // Ждем немного и проверяем, отправилась ли форма
        await sleep(2000);
        
        // Проверяем, есть ли сообщение об успешной отправке
        const successMessage = await page.evaluate(() => {
          const messages = Array.from(document.querySelectorAll('div, span'));
          return messages.find(msg => 
            msg.innerText && (
              msg.innerText.includes('Ваш ответ записан') ||
              msg.innerText.includes('Спасибо') ||
              msg.innerText.includes('ответ отправлен') ||
              msg.innerText.includes('успешно') ||
              msg.innerText.includes('записано') ||
              msg.innerText.includes('получен')
            )
          );
        });
        
        if (successMessage) {
          console.log('✅ Форма успешно отправлена!');
        } else {
          console.log('⚠️ Форма может не отправиться, пробую альтернативный способ...');
          
          // Проверяем, есть ли диалог подтверждения
          const confirmDialog = await page.evaluate(() => {
            const dialogs = Array.from(document.querySelectorAll('div[role="dialog"], div[role="alert"]'));
            return dialogs.find(dialog => 
              dialog.innerText && (
                dialog.innerText.includes('Отправить') ||
                dialog.innerText.includes('Подтвердить') ||
                dialog.innerText.includes('Submit')
              )
            );
          });
          
          if (confirmDialog) {
            console.log('✅ Найден диалог подтверждения, кликаю...');
            await confirmDialog.click();
            await sleep(2000);
          }
          
          // Способ 2: Клик через JavaScript
          await page.evaluate((btn) => {
            btn.dispatchEvent(new MouseEvent('click', {
              bubbles: true,
              cancelable: true,
              view: window
            }));
          }, submitBtn);
          
          await sleep(2000);
          
          // Способ 3: Нажатие Enter на кнопке
          await page.keyboard.press('Tab');
          await sleep(500);
          await page.keyboard.press('Enter');
          
          console.log('✅ Альтернативные способы отправки выполнены');
          
          // Финальная проверка отправки
          await sleep(3000);
          const finalCheck = await page.evaluate(() => {
            // Проверяем, изменился ли URL (признак успешной отправки)
            if (window.location.href.includes('formResponse')) {
              return 'URL изменился - форма отправлена';
            }
            
            // Проверяем сообщения об успехе
            const messages = Array.from(document.querySelectorAll('div, span'));
            const successMsg = messages.find(msg => 
              msg.innerText && (
                msg.innerText.includes('Ваш ответ записан') ||
                msg.innerText.includes('Спасибо') ||
                msg.innerText.includes('ответ отправлен') ||
                msg.innerText.includes('успешно') ||
                msg.innerText.includes('записано') ||
                msg.innerText.includes('получен')
              )
            );
            
            if (successMsg) {
              return 'Найдено сообщение об успехе';
            }
            
            // Проверяем, исчезла ли кнопка отправки
            const submitBtn = document.querySelector('div[aria-label="Submit"]');
            if (!submitBtn) {
              return 'Кнопка отправки исчезла - форма отправлена';
            }
            
            return 'Не удалось определить статус отправки';
          });
          
          console.log(`📊 Статус отправки: ${finalCheck}`);
          
          // Если форма не отправилась, попробуем еще раз
          if (finalCheck === 'Не удалось определить статус отправки') {
            console.log('🔄 Пробую повторную отправку...');
            
            // Ждем немного и пробуем еще раз
            await sleep(2000);
            
            // Находим кнопку заново
            const retryBtn = await page.evaluateHandle(() => {
              const btn = document.querySelector('div[aria-label="Submit"]');
              return btn;
            });
            
            const isRetryBtnDefined = await page.evaluate(btn => !!btn, retryBtn);
            if (isRetryBtnDefined) {
              // Пробуем клик через JavaScript с полными событиями
              await page.evaluate((btn) => {
                // Создаем полное событие клика
                const clickEvent = new MouseEvent('click', {
                  view: window,
                  bubbles: true,
                  cancelable: true,
                  clientX: btn.getBoundingClientRect().left + 10,
                  clientY: btn.getBoundingClientRect().top + 10
                });
                
                btn.dispatchEvent(clickEvent);
                
                // Также пробуем focus и enter
                btn.focus();
                setTimeout(() => {
                  btn.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter' }));
                  btn.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter' }));
                }, 100);
              }, retryBtn);
              
              await sleep(3000);
              
              // Финальная проверка
              const retryCheck = await page.evaluate(() => {
                if (window.location.href.includes('formResponse')) {
                  return 'URL изменился после повторной попытки';
                }
                
                const messages = Array.from(document.querySelectorAll('div, span'));
                const successMsg = messages.find(msg => 
                  msg.innerText && (
                    msg.innerText.includes('Ваш ответ записан') ||
                    msg.innerText.includes('Спасибо') ||
                    msg.innerText.includes('ответ отправлен')
                  )
                );
                
                if (successMsg) {
                  return 'Найдено сообщение об успехе после повторной попытки';
                }
                
                return 'Форма все еще не отправлена';
              });
              
              console.log(`📊 Статус после повторной попытки: ${retryCheck}`);
            }
          }
          
          // Ждем появления страницы "Ответ записан" и возвращаемся к форме
          console.log('⏳ Жду появления страницы "Ответ записан"...');
          await sleep(2000);
          
          // Проверяем, находимся ли мы на странице "Ответ записан"
          const isOnResponsePage = await page.evaluate(() => {
            // Проверяем URL
            if (window.location.href.includes('formResponse')) {
              return true;
            }
            
            // Проверяем текст на странице
            const pageText = document.body.innerText;
            if (pageText.includes('Ответ записан') || 
                pageText.includes('Ваш ответ записан') ||
                pageText.includes('Спасибо за ваш ответ')) {
              return true;
            }
            
            return false;
          });
          
          if (isOnResponsePage) {
            console.log('✅ Форма успешно отправлена! Найдена страница "Ответ записан"');
            
            // Возвращаемся к форме для следующего ответа
            console.log('🔄 Возвращаюсь к форме для следующего ответа...');
            await page.goto(FORM_URL, { waitUntil: 'networkidle2' });
            await sleep(2000);
            
            // Проверяем, что форма загрузилась
            const formLoaded = await page.evaluate(() => {
              const questions = document.querySelectorAll('div[role="listitem"]');
              return questions.length > 0;
            });
            
            if (formLoaded) {
              console.log('✅ Форма загружена, готов к следующему ответу');
            } else {
              console.log('⚠️ Форма не загрузилась, перезагружаю...');
              await page.reload({ waitUntil: 'networkidle2' });
              await sleep(2000);
            }
          } else {
            console.log('⚠️ Страница "Ответ записан" не найдена, форма может не отправиться');
          }
        }
        
      } catch (error) {
        console.log('❌ Ошибка при клике:', error.message);
        
        // Способ 4: Попробуем найти кнопку заново и кликнуть
        try {
          const newSubmitBtn = await page.$('div[role="button"]:has-text("Отправить")');
          if (newSubmitBtn) {
            await newSubmitBtn.click();
            console.log('✅ Повторный клик выполнен');
          }
        } catch (e) {
          console.log('❌ Повторный клик не удался:', e.message);
        }
      }
      
      await sleep(3000); // Ждем дольше после отправки
    } else {
      console.log('❗ Кнопка отправки не найдена!');
      
      // Попробуем найти кнопку по другому
      const alternativeBtn = await page.evaluate(() => {
        const buttons = Array.from(document.querySelectorAll('*'));
        return buttons.find(el => 
          el.innerText && el.innerText.toLowerCase().includes('отправить')
        );
      });
      
      if (alternativeBtn) {
        console.log('✅ Найдена альтернативная кнопка отправки');
        await alternativeBtn.click();
      }
    }
    
    // Ждем между отправками
    if (i < endRow - 1) {
      const delay = Math.floor(Math.random() * (DELAY_BETWEEN[1] - DELAY_BETWEEN[0] + 1)) + DELAY_BETWEEN[0];
      console.log(`⏳ Жду ${delay}мс перед следующей отправкой...`);
      await sleep(delay);
    }
  }

  await browser.close();
  console.log('\n🎉 Готово!');
}

main().catch(e => {
  console.error('❌ Ошибка:', e);
  process.exit(1);
}); 