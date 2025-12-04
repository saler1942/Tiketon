(function($) {
    'use strict';
    
    function updateUserInfo(selectElement) {
        var scannerId = selectElement.value;
        
        if (!scannerId) {
            // Очищаем поля, если сканер не выбран
            $('#id_first_name').val('');
            $('#id_last_name').val('');
            $('#id_email').val('');
            return;
        }
        
        // Получаем данные сканера через API
        fetch('/admin/scanner-info/' + scannerId + '/')
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    // Заполняем поля данными сканера
                    $('#id_first_name').val(data.first_name);
                    $('#id_last_name').val(data.last_name);
                    $('#id_email').val(data.email);
                    
                    // Показываем уведомление
                    showNotification('Данные сканера загружены автоматически');
                }
            })
            .catch(error => {
                console.error('Ошибка при загрузке данных сканера:', error);
                showNotification('Ошибка при загрузке данных сканера', 'error');
            });
    }
    
    function showNotification(message, type = 'success') {
        // Создаем уведомление
        var notification = $('<div>')
            .addClass('alert alert-' + type)
            .text(message)
            .css({
                'position': 'fixed',
                'top': '20px',
                'right': '20px',
                'z-index': '9999',
                'padding': '10px 15px',
                'border-radius': '4px',
                'background-color': type === 'success' ? '#d4edda' : '#f8d7da',
                'color': type === 'success' ? '#155724' : '#721c24',
                'border': '1px solid ' + (type === 'success' ? '#c3e6cb' : '#f5c6cb')
            });
        
        $('body').append(notification);
        
        // Удаляем уведомление через 3 секунды
        setTimeout(function() {
            notification.fadeOut(function() {
                notification.remove();
            });
        }, 3000);
    }
    
    // Делаем функцию доступной глобально
    window.updateUserInfo = updateUserInfo;
    
    // Инициализация при загрузке страницы
    $(document).ready(function() {
        // Добавляем обработчик изменения для поля сканера
        $('#id_scanner_admin').on('change', function() {
            updateUserInfo(this);
        });
    });
    
})(django.jQuery || jQuery);
