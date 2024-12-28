
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

%% Parametre optimizasyonu (grid search)
alpha_values = 0.1:0.1:1; % Alpha için aralık (0.1'den başlayarak 1'e kadar)
lambda_values = logspace(-3, 1, 20); % Lambda için aralık
bestR2 = -Inf; % En iyi R² başlangıç değeri
bestAlpha = 0;
bestLambda = 0;

for alpha = alpha_values
    for lambda = lambda_values
        predictedValues = zeros(size(outputValues_Test));
        
        % Her çıktı değişkeni için ayrı model eğitimi
        for i = 1:size(outputValues_Teaching, 2)
            y = outputValues_Teaching(:, i); % Hedef değişken
            [B, FitInfo] = lasso(inputValues_Teaching, y, 'Alpha', alpha, 'Lambda', lambda);
            predictedValues(:, i) = inputValues_Test * B + FitInfo.Intercept;
        end
        
        % R² hesapla
        R2_values = zeros(1, size(outputValues_Test, 2));
        for i = 1:size(outputValues_Test, 2)
            actual = outputValues_Test(:, i);
            predicted = predictedValues(:, i);
            R2_values(i) = calculateR2(actual, predicted);
        end
        
        % Ortalama R² hesapla
        meanR2 = mean(R2_values);
        if meanR2 > bestR2
            bestR2 = meanR2;
            bestAlpha = alpha;
            bestLambda = lambda;
        end
    end
end

disp('En iyi parametreler:');
disp(['Alpha: ', num2str(bestAlpha)]);
disp(['Lambda: ', num2str(bestLambda)]);
disp(['En iyi ortalama R²: ', num2str(bestR2)]);

%% En iyi parametrelerle modeli yeniden eğit
predictedValues = zeros(size(outputValues_Test));

for i = 1:size(outputValues_Teaching, 2)
    y = outputValues_Teaching(:, i); % Hedef değişken
    [B, FitInfo] = lasso(inputValues_Teaching, y, 'Alpha', bestAlpha, 'Lambda', bestLambda);
    predictedValues(:, i) = inputValues_Test * B + FitInfo.Intercept;
end

%% Hataların Hesaplanması
numOutputs = size(outputValues_Test, 2);
EN_ErrorValues = zeros(4, numOutputs);

for i = 1:numOutputs
    actual = outputValues_Test(:, i);
    predicted = predictedValues(:, i);
    EN_ErrorValues(1, i) = calculateR2(actual, predicted);
    EN_ErrorValues(2, i) = calculateMAPE(actual, predicted);
    EN_ErrorValues(3, i) = calculateMAE(actual, predicted);
    EN_ErrorValues(4, i) = calculateMSE(actual, predicted);
end

disp('--------------------------------------------');
disp('EN için Hata Değerleri (R², MAPE, MAE, MSE):');

% Hata değerlerini table formatında düzenle
errorMetrics = {'R2', 'MAPE', 'MAE', 'MSE'}'; % Hata türleri (sütun 1)
EN_ErrorValues_Table = array2table(EN_ErrorValues', ...
    'VariableNames', errorMetrics, ...
    'RowNames', {'Surface Roughness', 'Flank Wear', 'Cutting Temperature', 'Energy Consumption'});

% Tabloyu ekrana yazdır
disp(EN_ErrorValues_Table);

%% Grafik İşlemleri
figure('Name', 'ElasticNET Regression Tahmin Sonuçları');
tiledlayout(4, 1, 'TileSpacing', 'compact');

titles = {'Surface Roughness', 'Flank Wear', 'Cutting Temperature', 'Energy Consumption'};
numTests = size(inputValues_Test, 1);

for i = 1:4
    nexttile;
    plot(1:numTests, outputValues_Test(:, i), 'k-o', 'LineWidth', 1.5); hold on;
    plot(1:numTests, predictedValues(:, i), 'b-*', 'LineWidth', 1.5);
    title(titles{i});
    legend('Gerçek Değerler', 'Tahmin Edilen Değerler (EN)');
    grid on; hold off;
end

%% excel'e kaydet
algorithmName = 'ElasticNET'; % Algoritmanın ismi
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