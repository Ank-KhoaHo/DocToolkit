using System.Net;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Near misses on the IPv6 transition prefixes: addresses that match <b>part</b> of a prefix and
/// must NOT be unwrapped as one.
///
/// <c>GuardedResourceLoaderTests</c> proves the positive direction — a genuine 6to4, Teredo or
/// NAT64 wrapper around a private address is blocked. That is one half of the property. The other
/// half is that the byte tests are <i>conjunctions</i>: every byte must match. Mutation testing
/// found that half untested, surviving <c>&amp;&amp;</c> → <c>||</c> on all five prefix checks in
/// <c>TryGetEmbeddedIPv4</c>.
///
/// Each address below is a global-unicast address that is legitimately reachable today. Under a
/// loosened conjunction it would instead be unwrapped, reveal 10.0.0.5, and be refused — so a
/// caller would silently lose access to real hosts. The failure a loosened guard causes is
/// over-blocking, not under-blocking, which is why nothing else notices it.
/// </summary>
public class Ipv6TransitionNearMissTests
{
    [Theory]
    // NAT64 well-known, 64:ff9b::/96 — requires bytes 4-11 to be zero. Here they are not, so this is
    // an ordinary address that happens to share the 64:ff9b prefix.
    [InlineData("64:ff9b:203:405:607:809:a00:5")]
    // NAT64 local-use, 64:ff9b:1::/48 — requires b[4]=0x00 AND b[5]=0x01. Here b[5] is 0x02.
    [InlineData("64:ff9b:2:a00:0:500::")]
    // 6to4, 2002::/16 — requires b[0]=0x20 AND b[1]=0x02. Here b[1] is 0x03.
    [InlineData("2003:a00:5::")]
    // Teredo, 2001::/32 — requires b[0..3] = 20 01 00 00. Here b[3] is 0x01.
    // Bytes 12-15 are 10.0.0.5 XOR 0xFF, so a loosened check decodes them straight to 10.0.0.5.
    [InlineData("2001:1::f5ff:fffa")]
    public void AnAddressMatchingOnlyPartOfATransitionPrefixIsNotUnwrapped(string address)
    {
        Assert.False(
            GuardedResourceLoader.IsBlockedAddress(IPAddress.Parse(address)),
            $"{address} matches only part of a transition prefix and must be treated as the "
            + "ordinary global-unicast address it is. Blocking it means a byte test was loosened "
            + "from AND to OR, and real hosts are now unreachable.");
    }

    // The positive counterpart, kept beside the near misses so the pair reads as one statement:
    // change one byte and the same address IS a wrapper around 10.0.0.5, and IS refused. Without
    // this, every assertion above would still pass against a guard that unwrapped nothing at all.
    [Theory]
    [InlineData("64:ff9b::a00:5")]          // NAT64 well-known, bytes 4-11 zero
    [InlineData("64:ff9b:1:a00:0:500::")]   // NAT64 local-use, b[5] = 0x01
    [InlineData("2002:a00:5::")]            // 6to4, b[1] = 0x02
    [InlineData("2001::f5ff:fffa")]         // Teredo, b[3] = 0x00
    public void TheSameAddressWithTheFullPrefixIsBlocked(string address)
    {
        Assert.True(
            GuardedResourceLoader.IsBlockedAddress(IPAddress.Parse(address)),
            $"{address} is a transition wrapper around 10.0.0.5 and must be refused. If this "
            + "passes while the near-miss tests also pass, the guard is unwrapping nothing.");
    }

    /// <summary>
    /// The same property, on the bytes the theory above cannot reach.
    ///
    /// Those near misses all differ from a real prefix in a <b>late</b> byte, which only tests the
    /// last conjunct. Mutation testing measured the gap: eleven <c>&amp;&amp;</c> → <c>||</c>
    /// mutants at earlier positions survived the whole suite, because nothing fed the guard an
    /// address whose <i>first</i> byte is wrong while the rest of the prefix matches.
    ///
    /// Written as bytes rather than as <c>IPAddress.Parse</c> strings on purpose: which byte is
    /// wrong is the entire point of each case, and "2000:ff9b::a00:1" hides it behind the
    /// zero-compression rules.
    /// </summary>
    [Theory]
    // ---- NAT64 well-known, 64:ff9b::/96 - prefix is 00 64 FF 9B, bytes 4-11 zero ---------------
    [InlineData("2064ff9b" + "00000000" + "00000000" + "0a000001", "b[0]")]
    [InlineData("2000ff9b" + "00000000" + "00000000" + "0a000001", "b[1]")]
    [InlineData("2000009b" + "00000000" + "00000000" + "0a000001", "b[1] and b[2]")]
    // ---- NAT64 local-use, 64:ff9b:1::/48 - prefix is 00 64 FF 9B 00 01, v4 at b[6,7,9,10] ------
    [InlineData("2064ff9b" + "00010a00" + "00000100" + "00000000", "b[0]")]
    [InlineData("2000ff9b" + "00010a00" + "00000100" + "00000000", "b[1]")]
    [InlineData("2000009b" + "00010a00" + "00000100" + "00000000", "b[1] and b[2]")]
    [InlineData("20000000" + "00010a00" + "00000100" + "00000000", "b[1] through b[3]")]
    // ---- Teredo, 2001::/32 - prefix is 20 01 00 00, v4 at b[12..15] XOR 0xFF -------------------
    [InlineData("30010000" + "00000000" + "00000000" + "f5fffffe", "b[0]")]
    [InlineData("30300000" + "00000000" + "00000000" + "f5fffffe", "b[0] and b[1]")]
    // ---- IPv4-translated (SIIT) - eight zero bytes, then FF FF 00 00, then the v4 --------------
    [InlineData("30000000" + "00000000" + "ffff0000" + "0a000001", "the zero run at b[0]")]
    [InlineData("00000000" + "00000000" + "00ff0000" + "0a000001", "b[8]")]
    public void AnAddressMatchingOnlyTheTailOfATransitionPrefixIsNotUnwrapped(
        string hex, string wrongByte)
    {
        var address = new IPAddress(Convert.FromHexString(hex));

        Assert.False(
            GuardedResourceLoader.IsBlockedAddress(address),
            $"{address} is an ordinary global-unicast address: {wrongByte} does not match the "
            + "transition prefix whose remaining bytes it happens to share. Blocking it means a "
            + "conjunct was loosened to OR, so the guard now unwraps 10.0.0.1 out of an address "
            + "that never carried it - and real hosts became unreachable.");
    }


    /// <summary>
    /// The unwrapper reads FOUR CONSECUTIVE bytes starting at the prefix's offset. Mutation shifted
    /// the second of them - <c>ipv6Bytes[offset + 1]</c> became <c>[offset - 1]</c> - and every
    /// existing test still passed, because they all embed <c>10.0.0.5</c>: the shifted read picks up
    /// <c>10.2.0.5</c>, which is still inside <c>10.0.0.0/8</c> and still refused. The guard was
    /// reading the wrong byte and the result happened to land in the same block.
    ///
    /// <c>172.16.0.5</c> does not forgive that. It sits in <c>172.16.0.0/12</c>, a range that is
    /// private only for second octets 16 through 31 - so a shifted read yields <c>172.2.0.5</c>,
    /// which is ordinary public space, and the wrapper stops being refused.
    /// </summary>
    [Fact]
    public void ATransitionWrapperIsUnwrappedFromExactlyTheRightBytes()
    {
        // 6to4: 2002::/16 with the embedded v4 at bytes 2-5, here AC 10 00 05 = 172.16.0.5.
        var wrapper = IPAddress.Parse("2002:ac10:5::");

        Assert.True(
            GuardedResourceLoader.IsBlockedAddress(wrapper),
            "2002:ac10:5:: wraps 172.16.0.5 and must be refused. If this fails, the unwrapper is "
            + "reading a byte either side of the embedded address - which lands outside "
            + "172.16.0.0/12 and turns a private target back into a reachable one.");
    }
}
