
%% Excel dosyasından verileri oku
try
    data = readtable('Experiments Dosyasinin Kopyasi.xlsx', 'VariableNamingRule', 'preserve');
catch ME
    error(['Dosya okunamadı: ', ME.message]);
end

%% Öğretilecek verilerin belirlenmesi
data = data(:,2:end);
dataRows = size(data(:,1),1);

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

%% Lasso Regularization

% Her çıktı değişkeni için ayrı bir model oluştur
numOutputs = size(outputValues_Teaching, 2); % Çıktı sütunlarının sayısı
predictedValues_Test = zeros(size(outputValues_Test)); % Tahmin sonuçları için matris

for i = 1:numOutputs
    % Şu anki hedef sütun
    currentOutput_Teaching = outputValues_Teaching(:, i);

    % Lasso regresyonunu öğretim verileri ile eğit
    [B, FitInfo] = lasso(inputValues_Teaching, currentOutput_Teaching, 'CV', 10);

    % En iyi lambda değerini seç
    bestLambdaIndex = FitInfo.IndexMinMSE;
    bestB = B(:, bestLambdaIndex);
    bestIntercept = FitInfo.Intercept(bestLambdaIndex);

    % Test verilerini tahmin et
    predictedLassoValues_Test(:, i) = inputValues_Test * bestB + bestIntercept;
end

%% Hata hesaplama
 saveErrorValues = [];

% Her çıktı değişkeni için hata hesaplarını yap
for i = 1:numOutputs
    actual = outputValues_Test(:, i);
    predicted = predictedLassoValues_Test(:, i);

    R2 = calculateR2(actual, predicted);
    MAPE = calculateMAPE(actual, predicted);
    MAE = calculateMAE(actual, predicted);
    MSE = calculateMSE(actual, predicted);

   
    saveErrorValues(1,i) = R2;
    saveErrorValues(2,i) = MAPE;
    saveErrorValues(3,i) = MAE;
    saveErrorValues(4,i) = MSE;
end

disp('--------------------------------------------');
disp('Lasso için Hata Değerleri (R², MAPE, MAE, MSE):');

% Hata değerlerini table formatında düzenle
errorMetrics = {'R2', 'MAPE', 'MAE', 'MSE'}'; % Hata türleri (sütun 1)
Lasso_ErrorValues_Table = array2table(saveErrorValues', ...
    'VariableNames', errorMetrics, ...
    'RowNames', {'Surface Roughness', 'Flank Wear', 'Cutting Temperature', 'Energy Consumption'});

% Tabloyu ekrana yazdır
disp(Lasso_ErrorValues_Table);

%% Grafik İşlemleri
figure('Name', 'Lasso Regression Tahmin Sonuçları');
tiledlayout(4, 1, 'TileSpacing', 'compact');

titles = {'Surface Roughness', 'Flank Wear', 'Cutting Temperature', 'Energy Consumption'};
numTests = size(inputValues_Test, 1);

for i = 1:4
    nexttile;
    plot(1:numTests, outputValues_Test(:, i), 'k-o', 'LineWidth', 1.5); hold on;
    plot(1:numTests, predictedLassoValues_Test(:, i), 'b-*', 'LineWidth', 1.5);
    title(titles{i});
    legend('Gerçek Değerler', 'Tahmin Edilen Değerler (Lasso)');
    grid on; hold off;
end

%% excel'e kaydet
algorithmName = 'Lasso_Regression'; % Algoritmanın ismi
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
        predictedLassoValues_Test(:, outputIdx), ...  % Burada predictions olarak değiştirdik
        abs(outputValues_Test(:, outputIdx) - predictedLassoValues_Test(:, outputIdx)), ...  % predictions kullanıldı
        'VariableNames', {'Sample', 'Gerçek', 'Tahmin', 'Hata'});
    
    % Sayfa ismi belirleyelim
    sheetName = sheetNames{outputIdx};
    
    % Tabloyu Excel'e yaz
    writetable(comparisonTable, outputFileName, 'Sheet', sheetName);

    % Hata metriklerini hesapla ve yaz
    errorMetrics = {
        'R2', calculateR2(outputValues_Test(:, outputIdx), predictedLassoValues_Test(:, outputIdx));  % predictions kullanıldı
        'MAPE', calculateMAPE(outputValues_Test(:, outputIdx), predictedLassoValues_Test(:, outputIdx));  % predictions kullanıldı
        'MAE', calculateMAE(outputValues_Test(:, outputIdx), predictedLassoValues_Test(:, outputIdx));  % predictions kullanıldı
        'MSE', calculateMSE(outputValues_Test(:, outputIdx), predictedLassoValues_Test(:, outputIdx));  % predictions kullanıldı
    };
    errorTable = cell2table(errorMetrics, 'VariableNames', {'Hata_Metrikleri', 'Degerler'});
    
    % Hata metriklerini ilgili sayfaya yaz
    writetable(errorTable, outputFileName, 'Sheet', sheetName, 'Range', 'F1');
end

disp(['Sonuçlar Excel dosyasına kaydedildi: ', outputFileName]);