'use strict';

const path              = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');

const ROOT      = path.resolve(__dirname, '..');
const STUB_SCSS = path.resolve(__dirname, 'spfabric-stub.scss');

// Resolve a loader by absolute path (bypasses any loader-resolution ambiguity).
const r = name => require.resolve(name, { paths: [path.join(ROOT, 'node_modules')] });

// All @microsoft/sp-* packages need to be stubbed — their internals reference
// @ms/* packages that only exist inside the SPFx runtime.
const SP_PACKAGES = [
  '@microsoft/sp-webpart-base',
  '@microsoft/sp-property-pane',
  '@microsoft/sp-http',
  '@microsoft/sp-http-base',
  '@microsoft/sp-http-msgraph',
  '@microsoft/sp-core-library',
  '@microsoft/sp-component-base',
  '@microsoft/sp-lodash-subset',
  '@microsoft/sp-office-ui-fabric-core',
  '@microsoft/sp-page-context',
  '@microsoft/sp-adaptive-card-extension-base',
];

const spAliases = SP_PACKAGES.reduce((acc, pkg) => {
  acc[pkg] = pkg === '@microsoft/sp-webpart-base'
    ? path.resolve(__dirname, 'stubs/sp-webpart-base.js')
    : path.resolve(__dirname, 'stubs/sp-stub.js');
  return acc;
}, {});

// SPFx's real build resolves this module id via config.json's localizedResources
// loc-loader, which doesn't exist in this standalone demo bundle — point it
// straight at the English strings file instead.
spAliases['SmartOrgChartWebPartStrings'] = path.resolve(ROOT, 'src/webparts/smartOrgChart/loc/en-us.js');

module.exports = {
  mode:    'development',
  devtool: false,

  entry: path.resolve(__dirname, 'index.tsx'),

  output: {
    path:     path.resolve(__dirname, 'dist'),
    filename: 'bundle.js',
    clean:    true,
  },

  resolve: {
    extensions: ['.ts', '.tsx', '.js', '.jsx'],
    modules:    [path.join(ROOT, 'node_modules'), 'node_modules'],
    alias:      spAliases,
  },

  // Tell webpack where to find loaders.
  resolveLoader: {
    modules: [path.join(ROOT, 'node_modules'), 'node_modules'],
  },

  module: {
    rules: [
      // TypeScript / TSX
      {
        test:    /\.tsx?$/,
        exclude: /node_modules/,
        use: [{
          loader:  r('ts-loader'),
          options: {
            configFile:    path.resolve(__dirname, 'tsconfig.demo.json'),
            transpileOnly: true,
          },
        }],
      },

      // SCSS modules (*.module.scss)
      {
        test: /\.module\.scss$/,
        use: [
          r('style-loader'),
          {
            loader:  r('css-loader'),
            options: {
              modules: { localIdentName: '[name]__[local]' },
            },
          },
          {
            loader:  r('sass-loader'),
            options: {
              implementation: require(path.join(ROOT, 'node_modules/sass')),
              sassOptions: {
                importer(url) {
                  if (
                    url.includes('SPFabricCore') ||
                    url.includes('sp-office-ui-fabric-core')
                  ) {
                    return { file: STUB_SCSS };
                  }
                  return null;
                },
              },
            },
          },
        ],
      },

      // Plain SCSS (non-module)
      {
        test:    /\.scss$/,
        exclude: /\.module\.scss$/,
        use: [
          r('style-loader'),
          r('css-loader'),
          {
            loader:  r('sass-loader'),
            options: { implementation: require(path.join(ROOT, 'node_modules/sass')) },
          },
        ],
      },

      // Ignore resource files (.resx, .xml) that SPFx packages bundle
      {
        test: /\.(resx|xml|xlf)$/,
        type: 'asset/source',
      },
    ],
  },

  plugins: [
    new HtmlWebpackPlugin({
      template: path.resolve(__dirname, 'index.html'),
      inject:   'body',
    }),
    // Fluent UI icon glyphs — without these, every <Icon> in the app (toolbar
    // buttons, node card chevrons, etc.) renders as a blank square. SPFx
    // pulls this font from SharePoint's CDN; the demo harness has no network
    // access, so serve the same files locally instead.
    new CopyWebpackPlugin({
      patterns: [
        {
          from: path.join(ROOT, 'node_modules/@fluentui/font-icons-mdl2/fonts'),
          to:   path.resolve(__dirname, 'dist/fonts'),
        },
      ],
    }),
  ],
};
