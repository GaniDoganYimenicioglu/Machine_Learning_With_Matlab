
%% Excel dosyasından verileri oku
try
    data = readtable('Experiments Dosyasinin Kopyasi.xlsx', 'VariableNamingRule', 'preserve');
catch ME
    error(['Dosya okunamadı: ', ME.message]);
end

%% Öğretilecek ve test verilerinin belirlenmesi
data = data(:, 2:end); % İlk sütun çıkarıldı
dataRows = size(data, 1);

% [mmc type, cutting speed, feed rate, cooling/lubrication]
inputsValues_Teaching = data{1:130, 1:4};
outputValues_Teaching = data{1:130, 5:8};

% Test verileri
% [surface roughness, flank wear, cutting temperature, energy consumption]
inputsValues_Test = data{131:end, 1:4};
outputValues_Test = data{131:end, 5:8};

%% Hata hesaplama fonksiyonları
% R² hesaplama
calculateR2 = @(actual, predicted) 1 - sum((actual - predicted).^2) / sum((actual - mean(actual)).^2);

% MAPE hesaplama
calculateMAPE = @(actual, predicted) mean(abs((actual - predicted) ./ actual)) * 100;

% MAE hesaplama
calculateMAE = @(actual, predicted) mean(abs(actual - predicted));

% MSE hesaplama
calculateMSE = @(actual, predicted) mean((actual - predicted).^2);

%% Linear Regression modeli oluşturma ve eğitim
% Her bir çıktı değişkeni için ayrı modeller oluşturulacak
numOutputs = size(outputValues_Teaching, 2); % Çıktı sayısı
predictions = zeros(size(outputValues_Test)); % Test tahminleri için yer ayır

for i = 1:numOutputs
    % Mevcut çıktı için model oluştur
    mdl = fitlm(inputsValues_Teaching, outputValues_Teaching(:, i));
    
    % Test verileri ile tahmin yap
    predictions(:, i) = predict(mdl, inputsValues_Test);
end

%% Hata metriklerinin hesaplanması
% Her bir çıktı değişkeni için hata hesapla
saveErrorValues = [];

for i = 1:numOutputs
    actual = outputValues_Test(:, i);
    predicted = predictions(:, i);
    
    saveErrorValues(1,i) = calculateR2(actual, predicted);
    saveErrorValues(2,i) = calculateMAPE(actual, predicted);
    saveErrorValues(3,i) = calculateMAE(actual, predicted);
    saveErrorValues(4,i) = calculateMSE(actual, predicted);
end

disp('--------------------------------------------');
disp('Linear Regression için Hata Değerleri (R², MAPE, MAE, MSE):');

% Hata değerlerini table formatında düzenle
errorMetrics = {'R2', 'MAPE', 'MAE', 'MSE'}'; % Hata türleri (sütun 1)
LR_ErrorValues_Table = array2table(saveErrorValues', ...
    'VariableNames', errorMetrics, ...
    'RowNames', {'Surface Roughness', 'Flank Wear', 'Cutting Temperature', 'Energy Consumption'});

% Tabloyu ekrana yazdır
disp(LR_ErrorValues_Table);

%% Grafik Islemleri
figure('Name', 'Linear Regression Tahmin Sonuçları');
tiledlayout(4, 1, 'TileSpacing', 'compact');

titles = {'Surface Roughness', 'Flank Wear', 'Cutting Temperature', 'Energy Consumption'};
numTests = size(inputValues_Test, 1);

for i = 1:4
    nexttile;
    plot(1:numTests, outputValues_Test(:, i), 'k-o', 'LineWidth', 1.5); hold on;
    plot(1:numTests, predictions(:, i), 'b-*', 'LineWidth', 1.5);
    title(titles{i});
    legend('Gerçek Değerler', 'Tahmin Edilen Değerler (LR)');
    grid on; hold off;
end

%% excel'e kaydet
algorithmName = 'Linear_Regression'; % Algoritmanın ismi
outputFileName = [algorithmName, '_Results.xlsx'];

% Eğer dosya zaten varsa, sil
if isfile(outputFileName)
    delete(outputFileName);
end

% Sayfa isimleri
sheetNames = {'Surface Roughness', 'Flank Wear', 'Cutting Temperature', 'Energy Consumption'};

% Excel dosyasına yazılacak verilerin sırası
for outputIdx = 1:size(outputValues_Test, 2)
    % Gerçek ve tahmin edilen değerleri tabloya ekle
    comparisonTable = table((1:size(outputValues_Test, 1))', ...
        outputValues_Test(:, outputIdx), ...
        predictions(:, outputIdx), ...  % Burada predictions olarak değiştirdik
        abs(outputValues_Test(:, outputIdx) - predictions(:, outputIdx)), ...  % predictions kullanıldı
        'VariableNames', {'Sample', 'Gerçek', 'Tahmin', 'Hata'});
    
    % Sayfa ismi belirleyelim
    sheetName = sheetNames{outputIdx};
    
    % Tabloyu Excel'e yaz
    writetable(comparisonTable, outputFileName, 'Sheet', sheetName);

    % Hata metriklerini hesapla ve yaz
    errorMetrics = {
        'R2', calculateR2(outputValues_Test(:, outputIdx), predictions(:, outputIdx));  % predictions kullanıldı
        'MAPE', calculateMAPE(outputValues_Test(:, outputIdx), predictions(:, outputIdx));  % predictions kullanıldı
        'MAE', calculateMAE(outputValues_Test(:, outputIdx), predictions(:, outputIdx));  % predictions kullanıldı
        'MSE', calculateMSE(outputValues_Test(:, outputIdx), predictions(:, outputIdx));  % predictions kullanıldı
    };
    errorTable = cell2table(errorMetrics, 'VariableNames', {'Hata_Metrikleri', 'Degerler'});
    
    % Hata metriklerini ilgili sayfaya yaz
    writetable(errorTable, outputFileName, 'Sheet', sheetName, 'Range', 'F1');
end

disp(['Sonuçlar Excel dosyasına kaydedildi: ', outputFileName]);


