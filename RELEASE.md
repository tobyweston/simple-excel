# Releasing

## Prepare

In the `pom`, using HTTPS developer connection

``` xml
<scm>
    <developerConnection>scm:git:https://github.com/tobyweston/simple-excel.git</developerConnection>
</scm>
```

and running,

```
mvn release:prepare -Dusername=bub -Dpassword=secret
```

will update the `pom` and create a tag in SCM.


## Deploy

After running the `prepare`, run the `perform` step

```
mvn release:perform
```

to checkout the newly created tag, build to release and deploy it (using the `mvn deploy`). This will use the `distributionManagement` from the `pom`.

For local deployment to `robotooling.com`'s Maven repository, clone the repository before running this step.

```
git clone git@github.com:tobyweston/robotooling.git
git checkout gh-pages
...
perform the release process (inc deploy)
...
./update.sh
```


## Rollback

When things go wrong with the `prepare` step, rollback using

`mvn release:rollback -Dusername=bub -Dpassword=secret`

although you'll have to delete any tags manually

```
git tag -d simple-excel-1.x
git push origin :simple-excel-1.x
```