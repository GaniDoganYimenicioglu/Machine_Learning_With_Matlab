
%% Excel dosyasından verileri oku
try
    data = readtable('Experiments Dosyasinin Kopyasi.xlsx', 'VariableNamingRule', 'preserve');
catch ME
    error(['Dosya okunamadı: ', ME.message]);
end

%% Öğretilecek verilerin belirlenmesi
data = data(:, 2:end);
dataRows = size(data, 1);

% [mmc type, cutting speed, feed rate, cooling/lubrication]
inputValues_Teaching = data{1:130, 1:4};
outputValues_Teaching = data{1:130, 5:8};

%% Test verilerinin belirlenmesi
% [surface roughness, flank wear, cutting temperature, energy consumption]
inputValues_Test = data{130+1:dataRows, 1:4};
outputValues_Test = data{130+1:dataRows, 5:8};

%% Hata hesaplama fonksiyonları
% R² hesaplama
calculateR2 = @(actual, predicted) 1 - sum((actual - predicted).^2) / sum((actual - mean(actual)).^2);

% MAPE hesaplama
calculateMAPE = @(actual, predicted) mean(abs((actual - predicted) ./ actual)) * 100;

% MAE hesaplama
calculateMAE = @(actual, predicted) mean(abs(actual - predicted));

% MSE hesaplama
calculateMSE = @(actual, predicted) mean((actual - predicted).^2);

%% Random Forest modelinin eğitilmesi ve tahmin yapılması
numTrees = 100; % Orman içerisindeki ağaç sayısı
predictedValues = zeros(size(outputValues_Test));

for i = 1:size(outputValues_Teaching, 2)
    % Random Forest modelini oluştur
    rfModel = TreeBagger(numTrees, inputValues_Teaching, outputValues_Teaching(:, i), ...
                         'Method', 'regression', 'OOBPrediction', 'On');

    % Test verileri üzerinde tahmin yap
    predictedValues(:, i) = predict(rfModel, inputValues_Test);
end

%% Hata hesaplamaları
R2 = zeros(1, size(outputValues_Test, 2));
MAPE = zeros(1, size(outputValues_Test, 2));
MAE = zeros(1, size(outputValues_Test, 2));
MSE = zeros(1, size(outputValues_Test, 2));

for i = 1:size(outputValues_Test, 2)
    actual = outputValues_Test(:, i);
    predicted = predictedValues(:, i);

    % Hata metriklerini hesapla
    R2(i) = calculateR2(actual, predicted);
    MAPE(i) = calculateMAPE(actual, predicted);
    MAE(i) = calculateMAE(actual, predicted);
    MSE(i) = calculateMSE(actual, predicted);
end

%% Hata metriklerini görüntüle
disp('Hata Metrikleri:');
disp(table({'Surface Roughness'; 'Flank Wear'; 'Cutting Temperature'; 'Energy Consumption'}, ...
           R2', MAPE', MAE', MSE', ...
           'VariableNames', {'Output', 'R2', 'MAPE', 'MAE', 'MSE'}));

%% Grafik İşlemleri
figure('Name', 'Random Forest Tahmin Sonuçları');
tiledlayout(4, 1, 'TileSpacing', 'compact');

titles = {'Surface Roughness', 'Flank Wear', 'Cutting Temperature', 'Energy Consumption'};
numTests = size(inputValues_Test, 1);

for i = 1:4
    nexttile;
    plot(1:numTests, outputValues_Test(:, i), 'k-o', 'LineWidth', 1.5); hold on;
    plot(1:numTests, predictedValues(:, i), 'b-*', 'LineWidth', 1.5);
    title(titles{i});
    legend('Gerçek Değerler', 'Tahmin Edilen Değerler (RF)');
    grid on; hold off;
end

%% excel'e kaydet
algorithmName = 'Random_Forest'; % Algoritmanın ismi
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
        predictedValues(:, outputIdx), ...  % Burada predictions olarak değiştirdik
        abs(outputValues_Test(:, outputIdx) - predictedValues(:, outputIdx)), ...  % predictions kullanıldı
        'VariableNames', {'Sample', 'Gerçek', 'Tahmin', 'Hata'});
    
    % Sayfa ismi belirleyelim
    sheetName = sheetNames{outputIdx};
    
    % Tabloyu Excel'e yaz
    writetable(comparisonTable, outputFileName, 'Sheet', sheetName);

    % Hata metriklerini hesapla ve yaz
    errorMetrics = {
        'R2', calculateR2(outputValues_Test(:, outputIdx), predictedValues(:, outputIdx));  % predictions kullanıldı
        'MAPE', calculateMAPE(outputValues_Test(:, outputIdx), predictedValues(:, outputIdx));  % predictions kullanıldı
        'MAE', calculateMAE(outputValues_Test(:, outputIdx), predictedValues(:, outputIdx));  % predictions kullanıldı
        'MSE', calculateMSE(outputValues_Test(:, outputIdx), predictedValues(:, outputIdx));  % predictions kullanıldı
    };
    errorTable = cell2table(errorMetrics, 'VariableNames', {'Hata_Metrikleri', 'Degerler'});
    
    % Hata metriklerini ilgili sayfaya yaz
    writetable(errorTable, outputFileName, 'Sheet', sheetName, 'Range', 'F1');
end

disp(['Sonuçlar Excel dosyasına kaydedildi: ', outputFileName]);
