const gulp = require('gulp');
const zip = require('gulp-zip');
const clean = require('gulp-clean');
const uglify = require('gulp-uglify');
const concat = require('gulp-concat');

const paths = {
  scripts: {
    src: 'src/**/*.js',
    dest: 'dist/'
  },
  manifest: {
    src: 'manifest.json',
    dest: 'dist/'
  }
};

function cleanDist() {
  return gulp.src('dist', { allowEmpty: true, read: false })
    .pipe(clean());
}

function copyManifest() {
  return gulp.src(paths.manifest.src)
    .pipe(gulp.dest(paths.manifest.dest));
}

function scripts() {
  return gulp.src(paths.scripts.src)
    .pipe(uglify())
    .pipe(concat('app.min.js'))
    .pipe(gulp.dest(paths.scripts.dest));
}

function packageAddIn() {
  return gulp.src('dist/**/*')
    .pipe(zip('add-in.zip'))
    .pipe(gulp.dest('.'));
}

const build = gulp.series(cleanDist, gulp.parallel(copyManifest, scripts), packageAddIn);

exports.clean = cleanDist;
exports.copyManifest = copyManifest;
exports.scripts = scripts;
exports.package = packageAddIn;
exports.build = build;
