import React, { useState, useMemo } from 'react';
import { LineChart, Line, ScatterChart, Scatter, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, Label } from 'recharts';

const PoissonPowerLawInteraction = () => {
  const [lambda0, setLambda0] = useState(5);
  const [beta, setBeta] = useState(0.3);
  const [alpha, setAlpha] = useState(2.5);
  const [yMin, setYMin] = useState(100);

  // Poisson PMF
  const poissonPMF = (k, lambda) => {
    let logProb = k * Math.log(lambda) - lambda;
    for (let i = 1; i <= k; i++) {
      logProb -= Math.log(i);
    }
    return Math.exp(logProb);
  };

  // Power law PDF (continuous approximation)
  const powerLawPDF = (y, alpha, yMin) => {
    if (y < yMin) return 0;
    return (alpha - 1) / yMin * Math.pow(y / yMin, -alpha);
  };

  // Lambda as function of Y
  const lambdaFunction = (y) => {
    return lambda0 * Math.pow(y / yMin, beta);
  };

  // Generate power law distribution
  const powerLawData = useMemo(() => {
    const data = [];
    for (let y = yMin; y <= yMin * 50; y += yMin * 0.5) {
      data.push({
        y: y,
        pdf: powerLawPDF(y, alpha, yMin),
        lambda: lambdaFunction(y)
      });
    }
    return data;
  }, [alpha, yMin, lambda0, beta]);

  // Generate Poisson distributions for different Y values
  const poissonForDifferentY = useMemo(() => {
    const yValues = [yMin, yMin * 5, yMin * 20];
    const data = [];
    
    for (let k = 0; k <= 50; k++) {
      const point = { k };
      yValues.forEach(y => {
        const lambda = lambdaFunction(y);
        point[`Y=${y}`] = poissonPMF(k, lambda);
      });
      data.push(point);
    }
    return { data, yValues };
  }, [lambda0, beta, yMin]);

  // Generate joint distribution samples
  const jointSamples = useMemo(() => {
    const samples = [];
    // Generate Y values from power law (approximate sampling)
    for (let i = 0; i < 200; i++) {
      const u = Math.random();
      const y = yMin * Math.pow(1 - u, -1 / (alpha - 1));
      if (y > yMin * 100) continue; // Truncate extreme values
      
      const lambda = lambdaFunction(y);
      // Generate X from Poisson (approximate using normal for large lambda)
      let x;
      if (lambda > 20) {
        x = Math.round(lambda + Math.sqrt(lambda) * (Math.random() + Math.random() + Math.random() - 1.5));
      } else {
        // Simple Poisson sampling
        let L = Math.exp(-lambda);
        let p = 1;
        let k = 0;
        do {
          k++;
          p *= Math.random();
        } while (p > L);
        x = k - 1;
      }
      x = Math.max(0, x);
      
      samples.push({ y, x });
    }
    return samples;
  }, [lambda0, beta, alpha, yMin]);

  // Calculate summary statistics
  const stats = useMemo(() => {
    const example1Y = yMin * 10;
    const example2Y = yMin * 50;
    
    const lambda1 = lambdaFunction(example1Y);
    const lambda2 = lambdaFunction(example2Y);
    
    return {
      example1: {
        y: example1Y,
        lambda: lambda1,
        mean: lambda1,
        variance: lambda1,
        powerLawProb: powerLawPDF(example1Y, alpha, yMin)
      },
      example2: {
        y: example2Y,
        lambda: lambda2,
        mean: lambda2,
        variance: lambda2,
        powerLawProb: powerLawPDF(example2Y, alpha, yMin)
      }
    };
  }, [lambda0, beta, alpha, yMin]);

  return (
    <div className="w-full min-h-screen p-6 bg-gradient-to-br from-slate-50 to-indigo-50">
      <div className="max-w-7xl mx-auto space-y-6">
        <div className="bg-white rounded-lg shadow-lg p-6">
          <h1 className="text-3xl font-bold text-slate-800 mb-2">
            Poisson-Power Law Interaction Model
          </h1>
          <p className="text-slate-600 mb-4">
            Exploring how a power law variable modulates a Poisson process
          </p>
          
          <div className="bg-blue-50 border-l-4 border-blue-500 p-4 mb-4">
            <p className="font-mono text-sm">
              <strong>Model:</strong> λ(Y) = λ₀ × (Y/y_min)^β
            </p>
            <p className="font-mono text-sm">
              <strong>Joint:</strong> P(X,Y) = Poisson(X|λ(Y)) × PowerLaw(Y|α,y_min)
            </p>
          </div>

          {/* Controls */}
          <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">
                λ₀ (base rate): {lambda0}
              </label>
              <input
                type="range"
                min="1"
                max="20"
                step="1"
                value={lambda0}
                onChange={(e) => setLambda0(Number(e.target.value))}
                className="w-full"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">
                β (scaling exp): {beta.toFixed(2)}
              </label>
              <input
                type="range"
                min="0"
                max="1"
                step="0.05"
                value={beta}
                onChange={(e) => setBeta(Number(e.target.value))}
                className="w-full"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">
                α (tail exp): {alpha.toFixed(2)}
              </label>
              <input
                type="range"
                min="1.5"
                max="4"
                step="0.1"
                value={alpha}
                onChange={(e) => setAlpha(Number(e.target.value))}
                className="w-full"
              />
            </div>
            <div>
              <label className="block text-sm font-medium text-slate-700 mb-1">
                y_min: {yMin}
              </label>
              <input
                type="range"
                min="10"
                max="500"
                step="10"
                value={yMin}
                onChange={(e) => setYMin(Number(e.target.value))}
                className="w-full"
              />
            </div>
          </div>
        </div>

        {/* Power Law Distribution and Lambda Function */}
        <div className="grid md:grid-cols-2 gap-6">
          <div className="bg-white rounded-lg shadow-lg p-6">
            <h2 className="text-xl font-semibold text-slate-800 mb-4">
              Power Law Distribution (Factor Y)
            </h2>
            <ResponsiveContainer width="100%" height={250}>
              <LineChart data={powerLawData}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="y" scale="log" domain={['auto', 'auto']}>
                  <Label value="Y (log scale)" offset={-5} position="insideBottom" />
                </XAxis>
                <YAxis>
                  <Label value="P(Y)" angle={-90} position="insideLeft" />
                </YAxis>
                <Tooltip />
                <Line type="monotone" dataKey="pdf" stroke="#8b5cf6" strokeWidth={2} dot={false} />
              </LineChart>
            </ResponsiveContainer>
            <p className="text-sm text-slate-600 mt-2">
              P(Y) = (α-1)/y_min × (Y/y_min)^(-α)
            </p>
          </div>

          <div className="bg-white rounded-lg shadow-lg p-6">
            <h2 className="text-xl font-semibold text-slate-800 mb-4">
              Poisson Rate as Function of Y
            </h2>
            <ResponsiveContainer width="100%" height={250}>
              <LineChart data={powerLawData}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="y">
                  <Label value="Y" offset={-5} position="insideBottom" />
                </XAxis>
                <YAxis>
                  <Label value="λ(Y)" angle={-90} position="insideLeft" />
                </YAxis>
                <Tooltip />
                <Line type="monotone" dataKey="lambda" stroke="#059669" strokeWidth={2} dot={false} />
              </LineChart>
            </ResponsiveContainer>
            <p className="text-sm text-slate-600 mt-2">
              λ(Y) = λ₀ × (Y/y_min)^β = {lambda0} × (Y/{yMin})^{beta.toFixed(2)}
            </p>
          </div>
        </div>

        {/* Poisson Distributions */}
        <div className="bg-white rounded-lg shadow-lg p-6">
          <h2 className="text-xl font-semibold text-slate-800 mb-4">
            Conditional Poisson Distributions P(X|Y)
          </h2>
          <ResponsiveContainer width="100%" height={300}>
            <LineChart data={poissonForDifferentY.data}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="k">
                <Label value="X (count)" offset={-5} position="insideBottom" />
              </XAxis>
              <YAxis>
                <Label value="P(X|Y)" angle={-90} position="insideLeft" />
              </YAxis>
              <Tooltip />
              <Legend />
              <Line 
                type="monotone" 
                dataKey={`Y=${poissonForDifferentY.yValues[0]}`}
                stroke="#3b82f6" 
                strokeWidth={2} 
                dot={false} 
              />
              <Line 
                type="monotone" 
                dataKey={`Y=${poissonForDifferentY.yValues[1]}`}
                stroke="#f59e0b" 
                strokeWidth={2} 
                dot={false} 
              />
              <Line 
                type="monotone" 
                dataKey={`Y=${poissonForDifferentY.yValues[2]}`}
                stroke="#ef4444" 
                strokeWidth={2} 
                dot={false} 
              />
            </LineChart>
          </ResponsiveContainer>
          <p className="text-sm text-slate-600 mt-2">
            As Y increases, the Poisson distribution shifts right (higher mean) and spreads out
          </p>
        </div>

        {/* Joint Distribution Samples */}
        <div className="bg-white rounded-lg shadow-lg p-6">
          <h2 className="text-xl font-semibold text-slate-800 mb-4">
            Joint Distribution Samples P(X,Y)
          </h2>
          <ResponsiveContainer width="100%" height={350}>
            <ScatterChart>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="y" type="number">
                <Label value="Y (power law factor)" offset={-5} position="insideBottom" />
              </XAxis>
              <YAxis dataKey="x" type="number">
                <Label value="X (Poisson count)" angle={-90} position="insideLeft" />
              </YAxis>
              <Tooltip cursor={{ strokeDasharray: '3 3' }} />
              <Scatter data={jointSamples} fill="#8b5cf6" fillOpacity={0.6} />
            </ScatterChart>
          </ResponsiveContainer>
          <p className="text-sm text-slate-600 mt-2">
            Each point: (Y, X) sampled from joint distribution. Notice positive correlation due to λ(Y).
          </p>
        </div>

        {/* Example Calculations */}
        <div className="bg-white rounded-lg shadow-lg p-6">
          <h2 className="text-xl font-semibold text-slate-800 mb-4">
            Example Calculations
          </h2>
          
          <div className="grid md:grid-cols-2 gap-6">
            <div className="border-l-4 border-blue-500 pl-4">
              <h3 className="font-semibold text-lg mb-2">Example 1: Y = {stats.example1.y}</h3>
              <div className="space-y-1 text-sm font-mono">
                <p>λ(Y) = {lambda0} × ({stats.example1.y}/{yMin})^{beta.toFixed(2)}</p>
                <p className="text-blue-600 font-bold">λ = {stats.example1.lambda.toFixed(2)}</p>
                <p>E[X|Y] = {stats.example1.mean.toFixed(2)}</p>
                <p>Var[X|Y] = {stats.example1.variance.toFixed(2)}</p>
                <p className="text-purple-600 mt-2">P(Y) = {stats.example1.powerLawProb.toExponential(3)}</p>
              </div>
            </div>

            <div className="border-l-4 border-red-500 pl-4">
              <h3 className="font-semibold text-lg mb-2">Example 2: Y = {stats.example2.y}</h3>
              <div className="space-y-1 text-sm font-mono">
                <p>λ(Y) = {lambda0} × ({stats.example2.y}/{yMin})^{beta.toFixed(2)}</p>
                <p className="text-red-600 font-bold">λ = {stats.example2.lambda.toFixed(2)}</p>
                <p>E[X|Y] = {stats.example2.mean.toFixed(2)}</p>
                <p>Var[X|Y] = {stats.example2.variance.toFixed(2)}</p>
                <p className="text-purple-600 mt-2">P(Y) = {stats.example2.powerLawProb.toExponential(3)}</p>
              </div>
            </div>
          </div>

          <div className="mt-6 p-4 bg-slate-50 rounded">
            <h4 className="font-semibold mb-2">Key Insight:</h4>
            <p className="text-sm text-slate-700">
              As Y increases by {(stats.example2.y / stats.example1.y).toFixed(1)}×, the expected count E[X|Y] increases by{' '}
              {(stats.example2.lambda / stats.example1.lambda).toFixed(2)}×, but the probability P(Y) decreases by{' '}
              {(stats.example1.powerLawProb / stats.example2.powerLawProb).toFixed(1)}×. This creates a heavy-tailed count distribution with rare but extreme events.
            </p>
          </div>
        </div>
      </div>
    </div>
  );
};

export default PoissonPowerLawInteraction;