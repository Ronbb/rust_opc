pub trait TryIterator {
    type Item;
    type Error;

    fn try_next(&mut self) -> Result<Option<Self::Item>, Self::Error>;
}

pub struct TryIter<T: TryIterator> {
    inner: T,
    done: bool,
}

impl<T: TryIterator> Iterator for TryIter<T> {
    type Item = Result<T::Item, T::Error>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }

        match self.inner.try_next() {
            Ok(Some(item)) => Some(Ok(item)),
            Ok(None) => {
                self.done = true;
                None
            }
            Err(e) => {
                self.done = true;
                Some(Err(e))
            }
        }
    }
}

pub trait TryCacheIterator {
    type Item;
    type Error;
    type Cache: IntoIterator<Item = Self::Item>;

    fn try_cache(&mut self) -> Result<Option<Self::Cache>, Self::Error>;

    #[inline(always)]
    fn into_iter(self) -> TryCacheIter<Self>
    where
        Self: Sized,
    {
        TryCacheIter::new(self)
    }
}

pub struct TryCacheIter<T: TryCacheIterator> {
    inner: T,
    cache: Option<<<T as TryCacheIterator>::Cache as std::iter::IntoIterator>::IntoIter>,
    index: usize,
    done: bool,
}

impl<T: TryCacheIterator> TryCacheIter<T> {
    pub fn new(inner: T) -> Self {
        Self {
            inner,
            cache: None,
            index: 0,
            done: false,
        }
    }
}

impl<T: TryCacheIterator> Iterator for TryCacheIter<T> {
    type Item = Result<T::Item, T::Error>;

    fn next(&mut self) -> Option<Self::Item> {
        if self.done {
            return None;
        }

        let next = match &mut self.cache {
            Some(cache) => cache.next(),
            None => None,
        };

        match next {
            Some(item) => Some(Ok(item)),
            None => match self.inner.try_cache() {
                Ok(Some(cache)) => {
                    self.cache = Some(cache.into_iter());
                    self.index = 0;
                    self.next()
                }
                Ok(None) => {
                    self.done = true;
                    None
                }
                Err(e) => {
                    self.done = true;
                    Some(Err(e))
                }
            },
        }
    }
}
