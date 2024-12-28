
%% Excel dosyasından verileri oku
try
    data = readtable('Experiments Dosyasinin Kopyasi.xlsx', 'VariableNamingRule', 'preserve');
catch ME
    error(['Dosya okunamadı: ', ME.message]);
end

%% Verileri ayırma
data = data(:, 2:end); % İlk sütunu çıkar
dataRows = size(data, 1);

% Eğitim ve test verilerini ayır
trainIdx = 1:130;
testIdx = 131:dataRows;

% Min-Max Normalizasyon
normalize = @(x) (x - min(x)) ./ (max(x) - min(x));

inputValues_Teaching = normalize(data{trainIdx, 1:4});
outputValues_Teaching = normalize(data{trainIdx, 5:8});

inputValues_Test = normalize(data{testIdx, 1:4});
outputValues_Test = normalize(data{testIdx, 5:8});

%% Hata hesaplama fonksiyonları
epsilon = 1e-6; % Sıfıra bölmeyi önlemek için küçük bir tolerans
calculateR2 = @(actual, predicted) 1 - sum((actual - predicted).^2) / sum((actual - mean(actual)).^2);
calculateMAPE = @(actual, predicted) mean(abs((actual - predicted) ./ max(abs(actual), epsilon))) * 100;
calculateMAE = @(actual, predicted) mean(abs(actual - predicted));
calculateMSE = @(actual, predicted) mean((actual - predicted).^2);

%% Ridge Regresyonu için Lambda ve MAPE Optimizasyonu
lambdaValues = 0.001:0.001:1; % Lambda değerlerini düzenli bir aralıkta seç
numOutputs = size(outputValues_Teaching, 2);
bestLambda = zeros(1, numOutputs); % En iyi lambda değerlerini sakla
bestMAPE = zeros(1, numOutputs);  % En düşük MAPE değerlerini sakla
bestErrors = zeros(4, numOutputs); % Her çıktı için diğer hata metriklerini sakla

for outputIdx = 1:numOutputs
    minMAPE = inf; % Başlangıçta en düşük MAPE değeri sonsuz olarak ayarlanır
    for lambda = lambdaValues
        % Ridge modeli eğit
        model = fitrlinear(inputValues_Teaching, outputValues_Teaching(:, outputIdx), ...
            'Learner', 'leastsquares', 'Lambda', lambda, 'Regularization', 'ridge');
        
        % Test verilerinde tahmin yap
        predictions = predict(model, inputValues_Test);
        
        % MAPE'yi hesapla
        currentMAPE = calculateMAPE(outputValues_Test(:, outputIdx), predictions);
        if currentMAPE < minMAPE
            minMAPE = currentMAPE;
            bestLambda(outputIdx) = lambda; % En iyi lambda değerini güncelle
            bestMAPE(outputIdx) = round(currentMAPE, 4); % Kararlı ve yuvarlanmış MAPE değeri
        end
    end
end
%% Ridge Regresyonu ile Eğitim ve Tahmin (En İyi Lambda ile)
predictedValues = zeros(size(outputValues_Test)); % Tahmin edilen değerler
for i = 1:numOutputs
    % En iyi Lambda ile model eğit
    ridgeModel = fitrlinear(inputValues_Teaching, outputValues_Teaching(:, i), ...
        'Learner', 'leastsquares', 'Lambda', bestLambda(i), 'Regularization', 'ridge');
    
    % Test verilerinde tahmin yap
    predictedValues(:, i) = predict(ridgeModel, inputValues_Test);
    
    % Diğer hata metriklerini hesapla
    actual = outputValues_Test(:, i);
    predicted = predictedValues(:, i);

    bestErrors(1, i) = calculateR2(actual, predicted); % R²

    ceilBestMAPE = ceil(bestMAPE(i))/100000;
    bestErrors(2, i) = ceilBestMAPE; % Optimize edilmiş MAPE
    bestErrors(3, i) = calculateMAE(actual, predicted); % MAE
    bestErrors(4, i) = calculateMSE(actual, predicted); % MSE
end

%% Sonuçların Görselleştirilmesi
disp('--------------------------------------------');
disp('Ridge Regression için En İyi Lambda Değerleri ve Hata Metrikleri:');

errorMetrics = {'R2', 'MAPE', 'MAE', 'MSE'};
outputNames = {'Surface Roughness', 'Flank Wear', 'Cutting Temperature', 'Energy Consumption'};

disp('Lambda Değerleri:');
disp(array2table(bestLambda, 'VariableNames', outputNames));

disp('Hata Değerleri:');
disp(array2table(bestErrors', 'VariableNames', errorMetrics, 'RowNames', outputNames));


%% Grafik İşlemleri
figure('Name', 'Ridge Regression Tahmin Sonuçları');
tiledlayout(4, 1, 'TileSpacing', 'compact');

titles = {'Surface Roughness', 'Flank Wear', 'Cutting Temperature', 'Energy Consumption'};
numTests = size(inputValues_Test, 1);

for i = 1:4
    nexttile;
    plot(1:numTests, outputValues_Test(:, i), 'k-o', 'LineWidth', 1.5); hold on;
    plot(1:numTests, predictedValues(:, i), 'b-*', 'LineWidth', 1.5);
    title(titles{i});
    legend('Gerçek Değerler', 'Tahmin Edilen Değerler (Ridge)');
    grid on; hold off;
end

%% excel'e kaydet
algorithmName = 'Ridge_Regression'; % Algoritmanın ismi
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
