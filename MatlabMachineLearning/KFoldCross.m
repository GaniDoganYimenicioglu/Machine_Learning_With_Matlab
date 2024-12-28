
%% Excel dosyasından verileri oku
try
    data = readtable('Experiments Dosyasinin Kopyasi.xlsx', 'VariableNamingRule', 'preserve');
catch ME
    error(['Dosya okunamadı: ', ME.message]);
end

%% Verilerin ayrıştırılması
data = data(:,2:end);
dataRows = size(data(:,1),1);

% [mmc type, cutting speed, feed rate, cooling/lubrication]
inputValues_Teaching = data{1:130, 1:4};
outputValues_Teaching = data{1:130, 5:8};

% [surface roughness, flank wear, cutting temperature, energy consumption]
inputValues_Test = data{130+1:dataRows, 1:4};
outputValues_Test = data{130+1:dataRows, 5:8};

%% Hata hesaplama fonksiyonları
calculateR2 = @(actual, predicted) 1 - sum((actual - predicted).^2) / sum((actual - mean(actual)).^2);
calculateMAPE = @(actual, predicted) mean(abs((actual - predicted) ./ actual)) * 100;
calculateMAE = @(actual, predicted) mean(abs(actual - predicted));
calculateMSE = @(actual, predicted) mean((actual - predicted).^2);

%% K-Fold Cross Validation
k = 5; % K Fold sayısı
cv = cvpartition(size(inputValues_Teaching, 1), 'KFold', k);
hatalar = zeros(k, 4, size(outputValues_Teaching, 2)); % [R2, MAPE, MAE, MSE] hata değerleri her bir çıkış için

for outputIdx = 1:size(outputValues_Teaching, 2)
    for fold = 1:k
        % Verileri ayır
        trainIdx = training(cv, fold);
        testIdx = test(cv, fold);

        X_train = inputValues_Teaching(trainIdx, :);
        Y_train = outputValues_Teaching(trainIdx, outputIdx);
        X_test = inputValues_Teaching(testIdx, :);
        Y_test = outputValues_Teaching(testIdx, outputIdx);

        % Model oluştur ve eğit
        model = fitlm(X_train, Y_train);
        predictions = predict(model, X_test);

       
    end
end


%% Test verileriyle tahmin
predictedValues = zeros(size(outputValues_Test));
for outputIdx = 1:size(outputValues_Test, 2)
    finalModel = fitlm(inputValues_Teaching, outputValues_Teaching(:, outputIdx));
    predictedValues(:, outputIdx) = predict(finalModel, inputValues_Test);

    testR2(outputIdx) = calculateR2(outputValues_Test(:, outputIdx), predictedValues(:, outputIdx));
    testMAPE(outputIdx) = calculateMAPE(outputValues_Test(:, outputIdx), predictedValues(:, outputIdx));
    testMAE(outputIdx) = calculateMAE(outputValues_Test(:, outputIdx), predictedValues(:, outputIdx));
    testMSE(outputIdx) = calculateMSE(outputValues_Test(:, outputIdx), predictedValues(:, outputIdx));
end
%% Hata metriklerini görüntüle
disp('Hata Metrikleri:');
disp(table({'Surface Roughness'; 'Flank Wear'; 'Cutting Temperature'; 'Energy Consumption'}, ...
           testR2', testMAPE', testMAE', testMSE', ...
           'VariableNames', {'Output', 'R2', 'MAPE', 'MAE', 'MSE'}));

%% Grafik İşlemleri
figure('Name', 'K-Fold Cross');
tiledlayout(4, 1, 'TileSpacing', 'compact');

titles = {'Surface Roughness', 'Flank Wear', 'Cutting Temperature', 'Energy Consumption'};
numTests = size(inputValues_Test, 1);

for i = 1:4
    nexttile;
    plot(1:numTests, outputValues_Test(:, i), 'k-o', 'LineWidth', 1.5); hold on;
    plot(1:numTests, predictedValues(:, i), 'b-*', 'LineWidth', 1.5);
    title(titles{i});
    legend('Gerçek Değerler', 'Tahmin Edilen Değerler (KFC)');
    grid on; hold off;
end

%% excel'e kaydet
algorithmName = 'K_Fold_Cross'; % Algoritmanın ismi
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