'use strict';

const gulp = require('gulp');
const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

const subtaskbuildChild1 = build.subTask('sub-task-buildChild1', function (gulp, buildOptions, done) {

    console.log('sub-task-buildChild1 of build through console.log');
    this.log('sub-task-buildChild1 of build through this.log');
    
  
    
    this.logWarning('this is logWarning from sub-task-buildChild1');
    this.logError('this is logError from  sub-task-buildChild1');
  
    this.fileWarning('./dir1/subdir1/samplefile1.ts', 15, 20, 1520, 'warning code', 'this is fileWarning from sub-task-buildChild1');
    this.fileError('./dir1/subdir1/samplefile2.ts', 25, 20, 2520, 'error code', 'this is fileError from sub-task-buildChild1');
  
    
    done();

});
build.task('sub-task-buildChild1', subtaskbuildChild1);


const subtaskbuildChild2 = build.subTask('sub-task-buildChild2', function (gulp, buildOptions, done) {

    this.log('sub-task-buildChild2 of build through this.log');
    done();
    
});
build.task('sub-task-buildChild2', subtaskbuildChild2);


const postBundlesubTask = build.subTask('post-bundle', function (gulp, buildOptions, done) {
    this.log('Message from Post Bundle Task');
    done();
  });
  build.rig.addPostBundleTask(postBundlesubTask);



const preBuildSubTask = build.subTask('pre-build', function (gulp, buildOptions, done) {
    this.log('Message from PreBuild Task');
    done();
  });
  build.rig.addPreBuildTask(preBuildSubTask);

  const postBuildSubTask = build.subTask('post-build', function (gulp, buildOptions, done) {
    this.log('Message from PostBuild Task');
    done();
  });
  build.rig.addPostBuildTask(postBuildSubTask);

  const postTypeScriptSubTask = build.subTask('post-typescript', function (gulp, buildOptions, done) {
    this.log('Message from PostTypeScript task');
    done();
  });
  build.rig.addPostTypescriptTask(postTypeScriptSubTask);




build.initialize(gulp);

if (gulp.tasks['build']) {
    gulp.tasks['build'].dep.push('sub-task-buildChild1','sub-task-buildChild2');
  }

